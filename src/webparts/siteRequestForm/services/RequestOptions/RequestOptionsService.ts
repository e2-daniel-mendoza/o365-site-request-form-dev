import IRequestOptions from './IRequestOptions';

import { sp } from "@pnp/sp";

import { WebPartContext } from "@microsoft/sp-webpart-base";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class RequestOptionsService {
    private _optionsListTitle: string;

    constructor(optionsListTitle: string, webPartContext: WebPartContext) {
        this._optionsListTitle = optionsListTitle;
    }

    public async GetRequestOptions(): Promise<IRequestOptions[]> {
        const items = await sp.web.lists.getByTitle(this._optionsListTitle).items.select(
            "Title",
            "RestrictSiteCreation",
            "RestrictTemplateUsageToUsers/Id",
            "RestrictTemplateUsageToUsers/Name",
            "RequiresApproval",
            "Approvers/ID",
            "Approvers/Name",
            "AllowMSTeamsCreation",
            "UrlPrefix",
            "AddUserSetPrefix",
            "UserPrefixFormTitle",
            "SharePointTemplates/ID",
            "TeamsTemplates/ID"
        ).expand(
            "SharePointTemplates",
            "TeamsTemplates",
            "Approvers",
            "RestrictTemplateUsageToUsers"
        ).get();


        const options = items.map(i => {
            var approvers: string[] = [];
            if (i.Approvers != null && i.Approvers.length  > 0) {
                approvers = i.Approvers.map(subitem => { return subitem.Name; });
            }

            var restrictTemplateUsageToUsers: string[] = [];
            if (i.RestrictTemplateUsageToUsers != null && i.RestrictTemplateUsageToUsers.length > 0) {
                restrictTemplateUsageToUsers =  i.RestrictTemplateUsageToUsers.map(subitem => { return subitem.Name; });
            }

            return {
                Title: i.Title,
                RestrictSiteCreation: i.RestrictSiteCreation,
                RestrictSiteCreationToUserUids: restrictTemplateUsageToUsers,
                RequiresApproval: (i.RequiresApproval == null? false: i.RequiresApproval),
                ApproversUids: approvers,
                AllowMSTeamsCreation: i.AllowMSTeamsCreation,
                UrlPrefix: i.UrlPrefix,
                AddUserSetPrefix: (i.AddUserSetPrefix == null? false: i.AddUserSetPrefix),
                UserPrefixFormTitle: i.UserPrefixFormTitle,
                SharePointTemplatesIds: i.SharePointTemplates.map(spt => { return spt.ID; }),
                TeamsTemplatesIds: i.TeamsTemplates.map(tt => { return tt.ID; })
            };
        });

        return options;
    }
}