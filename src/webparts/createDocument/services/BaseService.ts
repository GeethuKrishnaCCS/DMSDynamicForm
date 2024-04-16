import { WebPartContext } from '@microsoft/sp-webpart-base';
//import { Constant } from "../shared/Constants";
// import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

export class BaseService {
  // private _spfi: SPFI;
  private _sp: SPFI;
  constructor(context: WebPartContext,siteUrl: string) {
      // this._sp = getSP(context);
      this._sp = new SPFI(siteUrl).using(SPFx(context));
    }
  public getCurrentUser() {
    return this._sp.web.currentUser();
    }
   
    public getListItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items();
    }
    public getItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id)();
    }
    public updateItem(url: string, listname: string, id: number,Details: any) {
        this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).update(Details);
    }
    public DeleteItem(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/Lists/" + listname).items.getById(id).delete();
    }
    public getLibraryItems(url: string, listname: string): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items();
    }
    public getLibraryItemById(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.getById(id)();
    }
    public updateLibraryItem(url: string, listname: string, id: number,Details: any) {
        this._sp.web.getList(url + "/" + listname).items.getById(id).update(Details);
    }
    public DeleteLibraryItem(url: string, listname: string, id: number): Promise<any> {
        return this._sp.web.getList(url + "/" + listname).items.getById(id).delete();
    }
    
    
} 