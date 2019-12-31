import { SPHttpClient, } from '@microsoft/sp-http';
import { Context } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PageContext } from '@microsoft/sp-page-context';
export interface IArabicformwebpartProps {
  description: string;
  greetings: string;
  LanguageKey: boolean;
  ListName: string;
  Date: string;
  textFieldId: string;
  CurrentLanauge: string;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  siteurl: string,
  ItemGuid: string,
  loading: false,
  UploadedFilesArray: Array<string>[];
  ProjectName: string;
  AmountForcast: Number;
  


}
