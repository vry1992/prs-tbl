import { useState } from 'react';
import './App.css';
import { ExcelReader } from './excel-reader';

function App() {
  const [isDownloadDisasbled, setIsDowloadDisables] = useState(true);
  const handleFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsDowloadDisables(true);
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        if (evt?.target?.result) {
          const buffer = evt.target.result as ArrayBuffer;

          const fileData = new Uint8Array(buffer);
          ExcelReader.read(fileData);
          const isOk = ExcelReader.isDownloadawailable();
          setIsDowloadDisables(!isOk);
        }
      } catch (error) {
        console.error(error);
        alert('Помилка читання файлу');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <>
      <input
        type="file"
        id="file"
        onChange={handleFile}
        accept=".xlsx, .xlsm"
      />
      <br />
      <button disabled={isDownloadDisasbled} onClick={ExcelReader.file}>
        Завантажити
      </button>
    </>
  );
}

export default App;
