import { StatusBar } from 'expo-status-bar';
import { StyleSheet, Text, View, Button, TextInput, ScrollView, ActivityIndicator } from 'react-native';
import { useEffect, useState } from 'react';
import axios from 'axios';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';


export default function App() {
  const [questions, setQuestions] = useState([]);
  const [variables, setVariables] = useState([]);
  const [answers, setAnswers] = useState({});
  const [loading, setLoading] = useState(false);
  const [showFinalButtons, setShowFinalButtons] = useState(false);
  const [message, setMessage] = useState('');

  const API_URL = 'http://192.168.18.240:8000'; // Replace with your PC's IP

  useEffect(() => {
    fetchQuestions();
  }, []);

  const fetchQuestions = async () => {
    setLoading(true);
    try {
      const response = await axios.get(`${API_URL}/get-questions`);
      const { questions, variables, next } = response.data;
      setQuestions(questions);
      setVariables(variables);

      if (!questions.length && !next) {
        setShowFinalButtons(true);
      }
    } catch (error) {
      console.error(error);
      setMessage("Error loading questions");
    }
    setLoading(false);
  };

  const handleAnswerChange = (key, value) => {
    setAnswers({ ...answers, [key]: value });
  };

  const submitAnswers = async () => {
    setLoading(true);
    try {
      await axios.post(`${API_URL}/submit-answers`, { answers });
      setAnswers({});
      fetchQuestions();
    } catch (error) {
      console.error(error);
      setMessage("Error submitting answers");
    }
    setLoading(false);
  };

  const generateJson = async () => {
    setLoading(true);
    try {
      const response = await axios.post(`${API_URL}/generate-json`);
      setMessage(response.data.message);
    } catch (error) {
      console.error(error);
      setMessage("Error generating JSON");
    }
    setLoading(false);
  };

  const generateDocx = async () => {
    setLoading(true);
    try {
      // First, tell the backend to generate the document
      await axios.post(`${API_URL}/generate-docx`);

      // Now download the document
      const downloadUrl = `${API_URL}/download-docx`;
      const fileUri = FileSystem.documentDirectory + 'PC1_Output.docx';

      const { uri } = await FileSystem.downloadAsync(downloadUrl, fileUri);
      console.log('Document downloaded to:', uri);

      // Open the sharing dialog
      await Sharing.shareAsync(uri);

      setMessage('Document downloaded successfully!');
    } catch (error) {
      console.error(error);
      setMessage('Error downloading document');
    }
    setLoading(false);
  };

  const restartForm = async () => {
    setLoading(true);
    try {
      await axios.post(`${API_URL}/restart`); // <-- Create this endpoint in backend
      setAnswers({});
      setQuestions([]);
      setVariables([]);
      setShowFinalButtons(false);
      setMessage('');
      fetchQuestions(); // load first questions
    } catch (error) {
      console.error(error);
      setMessage("Error restarting form");
    }
    setLoading(false);
  };




  return (
    <ScrollView contentContainerStyle={styles.container}>
      <Text style={styles.heading}>PC-1 Form Generator</Text>

      {loading && <ActivityIndicator size="large" color="blue" />}

      {!showFinalButtons && questions.map((question, index) => (
        <View key={index} style={{ width: '100%', marginBottom: 20 }}>
          <TextInput
            style={styles.input}
            placeholder={question}
            value={answers[variables[index]] || ''}
            onChangeText={(value) => handleAnswerChange(variables[index], value)}
          />
        </View>
      ))}

      {!showFinalButtons && (
        <Button title="Next" onPress={submitAnswers} />
      )}

      {showFinalButtons && (
        <View style={styles.finalButtonsContainer}>
          <Button title="Generate JSON File" onPress={generateJson} />
          <View style={styles.buttonSpacing} />

          <Button title="Generate Document" onPress={generateDocx} />
          <View style={styles.buttonSpacing} />

          <Button title="Review Document" onPress={generateDocx} />
          <View style={styles.buttonSpacing} />

          <Button title="Restart" color="red" onPress={restartForm} />
        </View>
      )}


      <Text style={styles.output}>{message}</Text>

      <StatusBar style="auto" />
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: {
    flexGrow: 1,
    backgroundColor: '#f9f9f9',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 20,
  },
  heading: {
    fontSize: 20,
    fontWeight: 'bold',
    marginBottom: 25,
    color: '#333',
  },
  input: {
    height: 40,
    width: '100%',
    borderColor: 'gray',
    borderWidth: 1,
    marginBottom: 12,
    paddingHorizontal: 10,
    borderRadius: 5,
    backgroundColor: '#fff',
  },
  output: {
    marginTop: 30,
    textAlign: 'center',
    fontSize: 16,
  },
  finalButtonsContainer: {
  width: '100%',
  marginTop: 30,
  alignItems: 'center',
 },

  buttonSpacing: {
    marginVertical: 8,
  },

});
