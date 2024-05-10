import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFormData {
  Nom: string; 
  Prenom: string; 
  Birthday: Date; // Définir le type de Birthday comme Date

}

export interface IFormProps {
  context: WebPartContext;
}
