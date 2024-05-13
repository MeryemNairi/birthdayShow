import * as React from 'react';
import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { IFormProps, IFormData } from './IFormProps';

export const Forme: React.FC<IFormProps> = ({ context }) => {
  const [formEntries, setFormEntries] = useState<IFormData[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);

  useEffect(() => {
    fetchFormData();
  }, []);

  useEffect(() => {
    const interval = setInterval(() => {
      setCurrentIndex(prevIndex => (prevIndex + 3 < formEntries.length ? prevIndex + 3 : 0));
    }, 5000);

    return () => clearInterval(interval);
  }, [formEntries]);

  const fetchFormData = async () => {
    try {
      const formData = await readExcelFile();
      const filteredFormData = formData.filter(entry => isToday(entry.Birthday));
      setFormEntries(filteredFormData);
    } catch (error) {
      console.error('Error fetching form data:', error);
    }
  };

  const readExcelFile = async (): Promise<IFormData[]> => {
    try {
      const response = await fetch('https://cnexia.sharepoint.com/sites/CnexiaForEveryone/Shared%20Documents/query.xlsx');
      const data = await response.arrayBuffer();

      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

      const formData: IFormData[] = rawData
        .slice(1)
        .map((row: any) => ({
          Nom: row[0],
          Prenom: row[1],
          Birthday: parseDate(row[2]) as Date,
        }));

      return formData;
    } catch (error) {
      console.error('Error reading Excel file:', error);
      throw new Error('An error occurred while reading Excel file. Please try again.');
    }
  };

  const parseDate = (dateString: any): Date | string => {
    if (!dateString) return '';

    if (typeof dateString === 'number') {
      return new Date((dateString - (25567 + 1)) * 86400 * 1000);
    } else if (typeof dateString === 'string') {
      return new Date(dateString);
    } else {
      return dateString;
    }
  };

  const isToday = (someDate: Date): boolean => {
    const today = new Date();
    return someDate.getDate() === today.getDate() &&
      someDate.getMonth() === today.getMonth();
  };

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center' }}>
        <h1 style={{ marginLeft: '10px', fontSize: '24px', fontWeight: 'bold', color: 'blue' }}>Birthday</h1>
      </div>
      {formEntries.slice(currentIndex, currentIndex + 3).map((entry, index) => (
        <div key={index} style={{ display: 'flex', marginBottom: '10px' }}>
          <div style={{ minWidth: '300px', border: '1px solid #ccc', borderRadius: '5px', padding: '10px', marginRight: '10px', backgroundColor: 'transparent' }}>
            <div style={{ fontSize: '18px', fontWeight: 'bold', marginBottom: '5px' }}>{entry.Nom} {entry.Prenom}</div>
            <div style={{ fontSize: '16px' }}>We wish you all a happy birthday!</div>
          </div>
        </div>
      ))}
    </div>
  );
};

export default Forme;
