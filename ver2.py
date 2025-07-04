import json
import os
from groq import Groq
from typing import Dict, List, Optional
from dataclasses import dataclass
from dotenv import load_dotenv
import sample2
import re

@dataclass
class State:
    name: str
    previous_state: Optional[str]
    condition: Optional[str]
    questions: List[str]
    variables: List[str]
    promptActions: List[str]
    promptFields: List[str]
    variableActions: List[str]
    next_state: Optional[str]

def parse_csv_line(line: str) -> List[str]:
    parts = []
    current = []
    in_array = False
    for char in line:
        if char == '[':
            in_array = True
        elif char == ']':
            in_array = False
        if char == ',' and not in_array:
            parts.append(''.join(current))
            current = []
        else:
            current.append(char)
    if current:
        parts.append(''.join(current))
    return parts

def parse_state_machine(csv_path: str) -> Dict[str, List[State]]:
    states = {}
    with open(csv_path, 'r', encoding='utf-8') as file:
        for line in file:
            if not line.strip():
                continue
            row = parse_csv_line(line.strip())
            state_name = row[0]
            prev_state = row[1] if row[1] != "null" else None
            condition = row[2] if row[2] != "null" else None
            try:
                questions = json.loads(row[3])
                variables = json.loads(row[4])
                actions = json.loads(row[5])
                pFields = json.loads(row[6])
                variable_actions = json.loads(row[7])
            except json.JSONDecodeError as e:
                print(f"Error parsing JSON in row: {row}\n{str(e)}")
                continue
            next_state = row[8] if row[8] != "null" else None
            state = State(
                name=state_name,
                previous_state=prev_state,
                condition=condition,
                questions=questions,
                variables=variables,
                promptActions=actions,
                promptFields=pFields,
                variableActions=variable_actions,
                next_state=next_state
            )
            if state_name not in states:
                states[state_name] = []
            states[state_name].append(state)
    return states

def replace_markers(text: str, user_data: Dict[str, str]) -> str:
    with open("prompts_with_json.json", 'r', encoding='utf-8') as file:
        json_data = json.load(file)

    def replace_match(match):
        var_name = match.group(1)
        value = user_data.get(var_name, f"UNKNOWN_{var_name}")
        return json.dumps(value, ensure_ascii=False) if isinstance(value, dict) else str(value)

    def replace_json_match(match):
        var_name = match.group(1)
        value = user_data.get(var_name, json_data.get(var_name, f"UNKNOWN_JSON_{var_name}"))
        return json.dumps(value, ensure_ascii=False)

    text = re.sub(r"@(\w+)", replace_json_match, text)
    text = re.sub(r"\^(\w+)", replace_match, text)

    return text

def evaluate_condition(condition: Optional[str], user_data: Dict[str, str]) -> bool:
    if not condition:
        return True

    value = user_data.get(condition.lstrip("!"), "").strip().lower()
    return value in ["yes", "true", "1"] if not condition.startswith("!") else value in ["no", "false", "0", "", None]

def clean_response(raw_response: str) -> str:
    cleaned = re.sub(r"<think>.*?</think>", "", raw_response, flags=re.DOTALL)
    cleaned = re.sub(r"(?i)(---json|```json|```)", "", cleaned).strip()
    return cleaned

def run_prompt_actions(state, user_data):
    from groq import Groq
    load_dotenv()
    api_key = os.getenv("GROQ_API_KEY")
    client = Groq(api_key=api_key)

    for i, action in enumerate(state.promptActions):
        prompt = replace_markers(action, user_data)
        field_name = state.promptFields[i] if i < len(state.promptFields) else None

        if not field_name:
            continue

        messages = [
            {"role": "system", "content": "You are a concise assistant..."},
            {"role": "user", "content": prompt}
        ]

        response_text = ""
        try:
            completion = client.chat.completions.create(
                model="deepseek-r1-distill-llama-70b",
                messages=messages,
                temperature=0.6,
                max_completion_tokens=4096,
                top_p=0.95,
                stream=True,
            )
            for chunk in completion:
                content = chunk.choices[0].delta.content or ""
                response_text += content

            cleaned_response = clean_response(response_text)
            try:
                user_data[field_name] = json.loads(cleaned_response)
            except json.JSONDecodeError:
                user_data[field_name] = cleaned_response

        except Exception as e:
            print(f"Error during completion: {str(e)}")

def run_state_machine(states):
    load_dotenv()
    api_key = os.getenv("GROQ_API_KEY")
    client = Groq(api_key=api_key)

    user_data = {}
    current_state = "q1"
    completed_states = set()

    while current_state:
        state_options = states.get(current_state, [])
        if not state_options:
            print(f"State '{current_state}' not found.")
            break

        valid_states = [s for s in state_options if evaluate_condition(s.condition, user_data)]
        if not valid_states:
            print(f"No valid state transitions found for state '{current_state}'")
            break

        selected_state = valid_states[0]

        for action in selected_state.variableActions:
            var, value = action.split("=")
            user_data[var] = None if value == "null" else value

        for question, variable in zip(selected_state.questions, selected_state.variables):
            user_input = input(f"{question}: ")
            user_data[variable] = user_input

        for i, action in enumerate(selected_state.promptActions):
            action = replace_markers(action, user_data)
            if i < len(selected_state.promptFields):
                field_name = selected_state.promptFields[i]
            else:
                print(f"Warning: No promptField defined for action '{action}', skipping.")
                continue

            messages = [
                {"role": "system", "content": "You are a concise assistant..."},
                {"role": "user", "content": action}
            ]

            response_text = ""
            try:
                completion = client.chat.completions.create(
                    model="deepseek-r1-distill-llama-70b",
                    messages=messages,
                    temperature=0.6,
                    max_completion_tokens=4096,
                    top_p=0.95,
                    stream=True,
                    stop=None,
                )
                for chunk in completion:
                    content = chunk.choices[0].delta.content or ""
                    response_text += content

                cleaned_response = clean_response(response_text)
                try:
                    user_data[field_name] = json.loads(cleaned_response)
                except json.JSONDecodeError:
                    user_data[field_name] = cleaned_response

            except Exception as e:
                print(f"Error during completion: {str(e)}")

        completed_states.add(current_state)
        current_state = selected_state.next_state

    print("✅ All required information has been gathered.")
    file_name = input("Enter filename for the document (without extension): ").strip() or "PC1_Wizard_Output"

    with open(f"{file_name}.json", "w", encoding="utf-8") as f:
        json.dump(user_data, f, ensure_ascii=False, indent=4)
    print(f"JSON saved as {file_name}.json")

    sample2.create_project_document_from_json(user_data, f"{file_name}.docx")
    print(f"Word document saved as {file_name}.docx")

def main():
    print("AIBee PC-1 AI (CLI Version)")
    csv_path = "state_machine.csv"
    if not os.path.exists(csv_path):
        print("State machine CSV file not found!")
        return
    states = parse_state_machine(csv_path)
    run_state_machine(states)

if __name__ == "__main__":
    main()

