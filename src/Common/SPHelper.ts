import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";


export default class SPHelper {


    public demoFunction = async () => {
        // let currentUser = await this.getCurrentUserInfo();
        // console.log(currentUser);
        // let azUserInfo = await graph.users.getById('revathy@o365practice.onmicrosoft.com').select('employeeId', 'displayName').get();
        // console.log(azUserInfo);
        let userToUpdate = await sp.web.siteUsers.getByEmail('revathy@o365practice.onmicrosoft.com').get();
        console.log(userToUpdate);
        await sp.profiles.setSingleValueProfileProperty(userToUpdate.LoginName, "Title", "Revathy Sudharsan");
        console.log("Updated");
    }

    public getCurrentUserInfo = async () => {
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