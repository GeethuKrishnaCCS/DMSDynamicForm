import { WebPartContext } from '@microsoft/sp-webpart-base';
//import { Constant } from "../shared/Constants";
// import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { getSP } from '../shared/PnP/pnpjsConfig';

export class BaseService {
  // private spfi: SPFI;
  private sp: SPFI;
  constructor(context: WebPartContext,siteUrl: string) {
      this.sp = getSP(context);
    //   this.sp = new SPFI(siteUrl).using(SPFx(context));
    }
  public getCurrentUser() {
    return this.sp.web.currentUser();
    }
   
    public getListItems(url: string, listname: string): Promise<any> {
        return this.sp.web.getList(url + "/Lists/" + listname).items();
    }
    public getItemById(url: string, listname: string, id: number): Promise<any> {
        return this.sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public updateItem(url: string, listname: string, id: number,Details: any) {
        this.sp.web.getList(url + "/Lists/" + listname).items.getById(id).update(Details);
    }
    public DeleteItem(url: string, listname: string, id: number): Promise<any> {
        return this.sp.web.getList(url + "/Lists/" + listname).items.getById(id).delete();
    }
    public getLibraryItems(url: string, listname: string): Promise<any> {
        return this.sp.web.getList(url + "/" + listname).items();
    }
    public getLibraryItemById(url: string, listname: string, id: number): Promise<any> {
        return this.sp.web.getList(url + "/" + listname).items.getById(id)();
    }
    public updateLibraryItem(url: string, listname: string, id: number,Details: any) {
        this.sp.web.getList(url + "/" + listname).items.getById(id).update(Details);
    }
    public DeleteLibraryItem(url: string, listname: string, id: number): Promise<any> {
        return this.sp.web.getList(url + "/" + listname).items.getById(id).delete();
    }
    public getCategoryItems(url: string, listname: string, departmentid:any): Promise<any> {
        return this.sp.web.getList(url + "/Lists/" + listname).items
            .filter("DepartmentId eq '" + departmentid + "'")
            .select("Department/ID, Department/Title,Title,ContentTypeName,Code,ID,Approver/Title,Approver/ID,Approver/EMail")
            .expand("Department,Approver")();
    }
    public getListGuid(url: string, List: string): Promise<any> {
        return this.sp.web.getList(url + "/" + List)()
    } 
    public async getContentTypeId(url: string, List: string): Promise<any> {
        return await this.sp.web.getList(url + "/" + List).contentTypes();
    }
    public EnsureUser(username: string) {
        return this.sp.web.ensureUser(username);
    }
    public createNewLibraryItem(url: string, listname: string, data: any): Promise<any> {
        return this.sp.web.getList(url + "/" + listname).items.add(data);
    }
    public createNewListItem(url: string, listname: string, data: any): Promise<any> {
        return this.sp.web.getList(url + "/Lists/" + listname).items.add(data);
    }
    public async uploadDocument(filename: string, filedata: any, libraryname: string): Promise<any> {
        const file = await this.sp.web.getFolderByServerRelativePath(libraryname)
            .files.addUsingPath(filename, filedata, { Overwrite: true });
            console.log('file: ', file);
        return file;
    }
} 