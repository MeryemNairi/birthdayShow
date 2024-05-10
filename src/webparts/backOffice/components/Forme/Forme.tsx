import * as React from 'react';
import * as XLSX from 'xlsx';
import { IFormProps, IFormData } from './IFormProps';

export const Forme: React.FC<IFormProps> = ({ context }) => {
  const [formEntries, setFormEntries] = React.useState<IFormData[]>([]);

  React.useEffect(() => {
    fetchFormData();
  }, []);

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

      // Récupérer les données brutes de la feuille Excel
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

      // Traitement des dates manuellement et conversion explicite du type
      const formData: IFormData[] = rawData
        .slice(1)
        .map((row: any) => ({
          Nom: row[0],
          Prenom: row[1],
          Birthday: parseDate(row[2]) as Date, // Conversion explicite du type en Date
        }));

      return formData;
    } catch (error) {
      console.error('Error reading Excel file:', error);
      throw new Error('An error occurred while reading Excel file. Please try again.');
    }
  };

  // Fonction pour analyser les dates
  const parseDate = (dateString: any): Date | string => {
    if (!dateString) return ''; // Gestion des cellules vides

    // Vérifier si la valeur est une date brute (nombre) ou une date en texte
    if (typeof dateString === 'number') {
      return new Date((dateString - (25567 + 1)) * 86400 * 1000); // Conversion du nombre en date
    } else if (typeof dateString === 'string') {
      return new Date(dateString); // La valeur est déjà une chaîne de caractères représentant une date
    } else {
      return dateString; // Retourner la valeur telle quelle
    }
  };

  // Fonction pour vérifier si une date est aujourd'hui
  const isToday = (someDate: Date): boolean => {
    const today = new Date();
    return someDate.getDate() === today.getDate() &&
           someDate.getMonth() === today.getMonth();
  };

  return (
    <div>
      <table>
        <thead>
          <tr>
            <th>Nom</th>
            <th>Prénom</th>
            <th>Birthday</th>
          </tr>
        </thead>
        <tbody>
          {formEntries.map((entry, index) => (
            <tr key={index}>
              <td>{entry.Nom}</td>
              <td>{entry.Prenom}</td>
              <td>{entry.Birthday.toLocaleDateString()}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default Forme;
