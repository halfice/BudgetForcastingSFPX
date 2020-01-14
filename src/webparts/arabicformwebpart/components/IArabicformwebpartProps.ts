import { SPHttpClient, } from '@microsoft/sp-http';
import { Context } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PageContext } from '@microsoft/sp-page-context';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';

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
  Screen:string;
  IsAuditorIsAdmin:boolean;
  Department:string;
  ProjectsArray:Array<object>;
  SelectedMonth:string;
  TotalAmountForcasted:string;
  MonthlyForcastAmount:string;
  BudgetForcasting:Array<object>;
  Remarks:string;
  MonthlyDeliveredAmount:string;
  BalanceForcastTotal:Number;
  BalanceDeliverTotal:Number;
  Monitoritems: Array<object>;
  MonitorColumns: IColumn[];
  MonitorIndex:number;
  showPanel: boolean;
  


}
