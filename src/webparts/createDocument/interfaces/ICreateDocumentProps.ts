import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICreateDocumentProps {
  context: WebPartContext;
  webpartHeader: string;
  siteUrl:string;
  department:string;
  category:string;
  sourceDocument:string;
 
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
  contentTypeName:string;
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
}