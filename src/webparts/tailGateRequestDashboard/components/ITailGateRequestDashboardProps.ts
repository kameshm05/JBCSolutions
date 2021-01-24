import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITailGateRequestDashboardProps {
  description: string;
  spcontext:WebPartContext;
  siteUrl:string;
}
