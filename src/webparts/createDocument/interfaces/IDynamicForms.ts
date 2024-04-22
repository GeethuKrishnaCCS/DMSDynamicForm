import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDynamicFormsProps {
    context: WebPartContext;
    siteUrl: string;
    contractIndex: string;
    contentTypeName: any;
    submitCallBack(callbackdata: any):any;
    saveCallBack(callbackdata: any):any;
    contractIndexId: any;
    listID: string;
    contentTypeId: string;
    disableDynamic: boolean;
    hideEdit: boolean;
    absolutesiteUrl: string;
}
export interface IDynamicFormsState {
    listID: string;
    contentTypeId: string;
    itemId: number;
    ChangedTitle: string;
    disableDynamic: boolean;
    hideEdit: boolean;
    cancelConfirmMsg: any;
    confirmDialog: boolean;
}