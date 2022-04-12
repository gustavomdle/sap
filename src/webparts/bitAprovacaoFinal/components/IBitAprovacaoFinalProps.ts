import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IBitAprovacaoFinalProps {
  description: string;
  siteurl: string;
  context: WebPartContext;
  statusBIT: string;
}
