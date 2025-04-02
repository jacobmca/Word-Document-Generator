import { useState } from 'react';
import PizZip from "pizzip";
import Docxtemplater from "docx-templater";
import { saveAs } from "file-saver";
import './App.css'

const WordTemplateFiller = () => {
  const [file, setFile] = useState(null);
  const [formData, setFormData] = useState({ name: "", date: "" });

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
  };

  const handleChange = (e) => {
    setFormData({ ...formData, [e.target.name]: e.target.value});
  }

  const generateDocument = async () => {
    if(!file) return alert("Upload a .doc file first");

    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = (event) => {
      const content = event.target.result;
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip);

      // Replace placeholders
      doc.setData(formData);
      doc.render();

      const blob = new Blob([doc.getZip().generate({ type: "arraybuffer"})], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      saveAs(blob, "filled-template.docx");
    };
  };

  return (
    <div>
      <h2>Fill Out Word Template</h2>
      <input type="file" accept=".docx" onChange={handleFileChange} />
      <input
        type="text"
        name="name"
        placeholder="Enter Name"
        onChange={handleChange}
      />
      <input type="date" name="date" onChange={handleChange} />
      <button onClick={generateDocument}>Generate Word File</button>
    </div>
  )
}

export default WordTemplateFiller;