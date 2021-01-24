import {  WebPartContext } from '@microsoft/sp-webpart-base';
export interface IReportSignOffProps {
  description: string;
  spcontext:WebPartContext;
  siteURL:string;
}
