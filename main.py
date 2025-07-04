# main.py (FastAPI Backend)

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import json
from dotenv import load_dotenv
import ver2

# Load environment variables
load_dotenv()

app = FastAPI()

# Allow requests from your Expo app
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Parse the state machine once at startup
csv_path = "state_machine.csv"
states = ver2.parse_state_machine(csv_path)

# Store user data & current state in-memory (or use a DB in production)
session_data = {
    "user_data": {},
    "current_state": "q1"
}

@app.get("/get-questions")
async def get_questions():
    """Return the questions for the current state."""
    current_state = session_data["current_state"]
    user_data = session_data["user_data"]

    state_options = states.get(current_state, [])
    valid_states = [s for s in state_options if ver2.evaluate_condition(s.condition, user_data)]
    if not valid_states:
        return {"questions": [], "variables": [], "next": None}

    selected_state = valid_states[0]
    return {
        "questions": selected_state.questions,
        "variables": selected_state.variables,
        "next": selected_state.next_state
    }

@app.post("/submit-answers")
async def submit_answers(request: Request):
    """Receive answers for the current state, run prompts if needed, and move to the next state."""
    data = await request.json()
    answers = data.get("answers", {})

    # Update user data with answers
    user_data = session_data["user_data"]
    user_data.update(answers)

    # Get the current state and process prompts
    current_state = session_data["current_state"]
    state_options = states.get(current_state, [])
    valid_states = [s for s in state_options if ver2.evaluate_condition(s.condition, user_data)]

    if not valid_states:
        return {"message": "No next state found.", "completed": True}

    selected_state = valid_states[0]

    # Run prompt actions (call Groq)
    ver2.run_prompt_actions(selected_state, user_data)

    # Move to the next state
    session_data["current_state"] = selected_state.next_state

    return {"message": "Answers saved", "next": session_data["current_state"]}

@app.post("/generate-json")
async def generate_json():
    """Save the user data as JSON."""
    file_name = "PC1_Output"
    user_data = session_data["user_data"]

    with open(f"{file_name}.json", "w", encoding="utf-8") as f:
        json.dump(user_data, f, ensure_ascii=False, indent=4)

    return {"message": "JSON file generated ✅", "filename": f"{file_name}.json"}

@app.post("/generate-docx")
async def generate_docx():
    """Generate the Word document using sample2."""
    file_name = "PC1_Output"
    user_data = session_data["user_data"]

    import sample2
    sample2.create_project_document_from_json(user_data, f"{file_name}.docx")

    return {"message": "Word document generated ✅", "filename": f"{file_name}.docx"}

@app.get("/download-docx")
def download_docx():
    file_path = "PC1_Output.docx"  # This should match your generate-docx filename
    if not os.path.exists(file_path):
        return {"error": "Document not found. Please generate the document first."}

    return FileResponse(
        file_path,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename="PC1_Output.docx"
    )

@app.post("/restart")
async def restart():
    global state_machine, user_data, current_state
    user_data = {}  # Clear all user data
    current_state = "q1"  # Reset to first state
    return {"message": "Form restarted"}
