import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBitTodosItensProps {
  description: string;
  siteurl: string;
  context: WebPartContext;
  statusBIT: string;
}
