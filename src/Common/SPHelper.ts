import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import "@pnp/sp/profiles";
import "@pnp/sp/search";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ISearchQuery, SearchResults, SearchQueryBuilder } from "@pnp/sp/search";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import * as moment from 'moment';

export interface IUserInfo {
    ID: number;
    Email: string;
    LoginName: string;
    DisplayName: string;
    Picture: string;
}

export interface IPropertyMappings {
    AzProperty: string;
    SPProperty: string;
}

export interface IPropertyPair {
    name: string;
    value: string;
}

export interface IUserPropertyMapping {
    UserID: string;
    Properties: IPropertyPair[];
}

export interface IJsonMapping {
    targetSiteUrl?: string;
    targetAdminUrl?: string;
    values?: IUserPropertyMapping[];
}

export interface ISPHelper {
    demoFunction: () => void;
    getCurrentUserInfo: () => Promise<IUserInfo>;
    getPropertyMappings: () => Promise<any>;
    addFilesToFolder: (filename: string, fileContent: any) => void;
}

export default class SPHelper implements ISPHelper {

    private SiteURL: string = "";
    private SiteRelativeURL: string = "";
    private AdminSiteURL: string = "";
    private SyncFilePath: string = "/Shared Documents/SyncJobTemplate/";
    private SyncFileName: string = `SyncTemplate_${moment().format("MM-DD-YYYY-HH-mm-ss")}.json`

    constructor(siteurl: string, tenantname: string, domainname: string, relativeurl: string) {
        this.SiteURL = siteurl;
        this.SiteRelativeURL = relativeurl;
        this.AdminSiteURL = `https://${tenantname}-admin.${domainname}`;
    }

    public demoFunction = async () => {
        // let currentUser = await this.getCurrentUserInfo();
        // console.log(currentUser);
        // let azUserInfo = await graph.users.getById('revathy@o365practice.onmicrosoft.com').select('employeeId', 'displayName').get();
        // console.log(azUserInfo);
        let userToUpdate = await sp.web.siteUsers.getByEmail('revathy@o365practice.onmicrosoft.com').get();
        console.log(userToUpdate);
        // await sp.profiles.setSingleValueProfileProperty(userToUpdate.LoginName, "Title", "Revathy Sudharsan");
        // console.log("Updated");



        const result = await sp.profiles.clientPeoplePickerSearchUser({
            AllowEmailAddresses: true,
            AllowMultipleEntities: false,
            MaximumEntitySuggestions: 25,
            QueryString: 'Manager:*sudha*'
        });
        console.log(result);

        const results2: SearchResults = await sp.search("Manager:*Sudha*");
        console.log(results2);
    }
    /**
     * Generated the property mapping json content.
     */
    public getPropertyMappings = async (): Promise<any> => {
        let propertyMappings: IPropertyMappings[] = await sp.web.lists.getByTitle('Sync Properties Mapping').items
            .select("AzProperty", "SPProperty")
            .filter(`IsActive eq 1`)
            .get();
        let finalJson: string = "";
        let propertyFair: IPropertyPair[] = [];
        propertyMappings.map((propsMap: IPropertyMappings) => {
            propertyFair.push({
                name: propsMap.SPProperty,
                value: ""
            });
        });
        finalJson = `{
            "targetAdminUrl": "${this.AdminSiteURL}",
            "targetSiteUrl": "${this.SiteURL}",
            "values": [
                {
                    "UserID": "userid@tenantname.onmicrosoft.com",
                    "Properties": ${JSON.stringify(propertyFair)}
                }
            ]
        }`;
        //console.log(finalJson);
        return JSON.parse(finalJson);
    }
    /**
     * Add a file to a folder with contents.
     * This is used for creating the template json file.
     */
    public addFilesToFolder = async (fileContent: any) => {
        await this.checkAndCreateSyncTemplateFolder();
        return await sp.web.getFolderByServerRelativeUrl(`${this.SiteRelativeURL}${this.SyncFilePath}`)
            .files
            .add(decodeURI(this.SiteRelativeURL+this.SyncFilePath+this.SyncFileName), fileContent, true);
    }
    /**
     * Check for the template folder, if not creates.
     */
    public checkAndCreateSyncTemplateFolder = async () => {
        try {
            await sp.web.getFolderByServerRelativeUrl(`${this.SiteRelativeURL}${this.SyncFilePath}`).get();
        } catch (err) {
            await sp.web.folders.add(`${this.SiteRelativeURL}${this.SyncFilePath}`);
        }
    }
    /**
     * Get current logged in user information.
     */
    public getCurrentUserInfo = async (): Promise<IUserInfo> => {
        let currentUserInfo = await sp.web.currentUser.get();
        return ({
            ID: currentUserInfo.Id,
            Email: currentUserInfo.Email,
            LoginName: currentUserInfo.LoginName,
            DisplayName: currentUserInfo.Title,
            Picture: '/_layouts/15/userphoto.aspx?size=S&username=' + currentUserInfo.UserPrincipalName,
        });
    }
}