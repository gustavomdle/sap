import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBitEditarItemProps {
  description: string;
  siteurl: string;
  context: WebPartContext;
  addUsersAprovadorEngenharia: string[];
  addUsersAprovadorGeral: string[];
  addUsersDestinatariosAdicionais: string[];
  defaultmyusers?: any[];
  _addUsersAprovadorEngenharia: [];
}
