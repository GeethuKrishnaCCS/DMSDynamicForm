import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICreateDocumentProps {
  context: WebPartContext;
  webpartHeader: string;
  siteUrl:string;
  absolutesiteUrl:string;
  department:string;
  category:string;
  documentIdSettings:string;
  documentIdSequenceSettings:string;
  sourceDocument:string;
  publishDocument:string;
 
}
export interface ICreateDocumentState {
  currentUserId:any;
  currentUserName: string;
  currentUserEmail: string;
  title : string;
  department:string;
  departmentId:any;
  departmentOption:any[];
  departmentCode:string;
  category:string;
  categoryId:any;
  categoryOption:any[];
  categoryCode:string;
  contentTypeArray:any[]
  contentTypeId:string;
  sourcecontentTypeId:string;
  publishcontentTypeId:string;
  contentTypeName:string;
  sourcelistID:string;
  publishlistID:string;
  listID:string;
  reviewers: any[];
  reviewersDetails: any[];
  reviewersEmail: any;
  reviewersName: any;
  approver: any;
  approverEmail: any;
  approverName: any;
  disableDynamic:boolean;
  dynamic:boolean;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  selectedKey: any;
  proceedButton: boolean;
  hideEdit: boolean;
  showSubmitBTN: boolean;
  hideSubmitBTN: boolean;
  DocumentId:any;
  mydoc: File | null;
  incrementSequenceNumber:string;
  documentid: string;
  documentName: string;
}