import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAddTailGateRequestProps {
  description: string;
  spcontext:WebPartContext;
  siteURL:string;
}
