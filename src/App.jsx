import { useState } from "react";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import "./App.css";

const WordTemplateFiller = () => {
  const [file, setFile] = useState(null);
  const [formData, setFormData] = useState({
    Date: '',
    Project: '',
    JobTitle: ''
  });

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleChange = (e) => {
    const { name, value } = e.target;

    if (name === 'Date') {
      // For the date field, store the original format
      setFormData(prevState => ({
        ...prevState,
        [name]: value // Keep the yyyy-MM-dd format for the input field
      }));
    } else {
      setFormData(prevState => ({
        ...prevState,
        [name]: value
      }));
    }
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
        let formattedDate = formData.Date;
        if (formData.Date) {
          const date = new Date(formData.Date);
          formattedDate = date.toLocaleDateString('en-US', {
            month: '2-digit',
            day: '2-digit',
            year: 'numeric'
          });
        }
  
        const templateData = {
          Date: formattedDate,
          Project: formData.Project,
          JobTitle: formData.JobTitle
        };
  
        console.log('Template Data:', JSON.stringify(templateData, null, 2));
  
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
          data: templateData
        });
  
        console.log('About to render document');
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
          name="Date"
          onChange={handleChange}
          value={formData.Date || ''}
        />
        <input
          type="text"
          name="Project"
          placeholder="Enter Project"
          onChange={handleChange}
          value={formData.Project || ''}
        />
        <input
          type="text"
          name="JobTitle"
          placeholder="Enter Job Title"
          onChange={handleChange}
          value={formData.JobTitle || ''}
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
