import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
// import { getSP } from "../shared/PnP/pnpjsConfig";
import { SPFI, SPFx } from "@pnp/sp";

export class DynamicFormsServices extends BaseService {
    private _spfi: SPFI;
    constructor(context: WebPartContext, siteUrl: string) {
        super(context, siteUrl);
        // this._spfi = getSP(context);
        this._spfi = new SPFI(siteUrl).using(SPFx(context));
    }
    public _getListGuid(url: string, List: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + List)()
    }
    public async _getContentTypeId(url: string, List: string): Promise<any> {
        return await this._spfi.web.getList(url + "/Lists/" + List).contentTypes();
    }
    public updateItem(url: string, listname: string, data: any, id: number): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public _getMandatory(url: string, List: string, contenttypeID: any): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + List)
            .contentTypes.getById(contenttypeID).fields.filter("Required eq true")();
    }
}
