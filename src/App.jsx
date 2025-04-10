import { useState } from "react";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import "./App.css";

const WordTemplateFiller = () => {
  const [file, setFile] = useState(null);
  const [formData, setFormData] = useState({
    date: '',
    artist: '',
    project: '',
    jobTitle: '',
    occupationCode: '',
    wageScale: '',
    hours: '',
    hourlyRate: ''
  });

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prevState => ({
      ...prevState,
      [name]: value
    }));
  };

  const generateDocument = async () => {
    if (!file) return alert("Upload a .doc file first");
  
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = (event) => {
      try {
        const content = event.target.result;
        console.log('File loaded successfully, content length:', content.byteLength);
        
        const zip = new PizZip(content);
        console.log('ZIP created successfully');
  
        // Format the date for the document
        const formattedDate = formData.date 
        ? new Date(formData.date).toLocaleDateString('en-US', {
            month: '2-digit',
            day: '2-digit',
            year: 'numeric'
          })
        : '';
  
        const templateData = {
          date: formattedDate,
          artist: formData.artist,
          project: formData.project,
          jobTitle: formData.jobTitle,
          occupationCode: formData.occupationCode,
          wageScale: formData.wageScale,
          hours: formData.hours,
          hourlyRate: formData.hourlyRate
        };
  
        console.log('Template Data:', JSON.stringify(templateData, null, 2));
  
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
          delimiters: {
            start: '{',  // Single curly brace
            end: '}'     // Single curly brace
          },
          debug: true,
          data: templateData
        });

        doc.setData(templateData);
    
        console.log('About to render document');
        console.log('Final template data being used:', templateData);
        doc.render();
        console.log('Document rendered successfully');
  
        const blob = new Blob(
          [doc.getZip().generate({ type: "arraybuffer" })],
          {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          }
        );
  
        saveAs(blob, "filled-template.docx");
      } catch (error) {
        // If the error contains additional information in the properties
        if (error.properties && error.properties.errors instanceof Array) {
          const errorMessages = error.properties.errors.map(error => error.properties.explanation).join("\n");
          console.error('Templating errors: ', errorMessages);
        }
        console.error("Detailed error:", {
          message: error.message,
          properties: error.properties,
          stack: error.stack
        });
        alert("Error generating document. Check console for details.");
      }
    };
  };  

  return (
    <div>
      <h2>Fill Out Word Template</h2>
      <input type="file" accept=".docx" onChange={handleFileChange} />
      <div>
        <input
          type="date"
          name="date"
          onChange={handleChange}
          value={formData.date || ''}
        />
        <input
          type="text"
          name="artist"
          placeholder="Enter Artist Name"
          onChange={handleChange}
          value={formData.artist || ''}
        />
        <input
          type="text"
          name="project"
          placeholder="Enter Project"
          onChange={handleChange}
          value={formData.project || ''}
        />
        <input
          type="text"
          name="jobTitle"
          placeholder="Enter Job Title"
          onChange={handleChange}
          value={formData.jobTitle || ''}
        />
        <input
          type="text"
          name="occupationCode"
          placeholder="Enter Occupation Code"
          onChange={handleChange}
          value={formData.occupationCode || ''}
        />
        <input
          type="text"
          name="wageScale"
          placeholder="Enter Wage Scale"
          onChange={handleChange}
          value={formData.wageScale || ''}
        />
        <input
          type="text"
          name="hours"
          placeholder="Enter Hours"
          onChange={handleChange}
          value={formData.hours || ''}
        />
        <input
          type="text"
          name="hourlyRate"
          placeholder="Enter Hourly Rate"
          onChange={handleChange}
          value={formData.hourlyRate || ''}
        />
      </div>
      <button onClick={generateDocument}>Generate Word File</button>
    </div>
  );
};

function App() {
  return (
    <div className="App">
      <WordTemplateFiller />
    </div>
  );
}

export default App;
