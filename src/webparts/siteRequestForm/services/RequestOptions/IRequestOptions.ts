export default interface IRequestOptions {
    Title: string;
    RestrictSiteCreation: boolean;
    RestrictSiteCreationToUserUids: string[];
    RequiresApproval: boolean;
    ApproversUids: string[];
    AllowMSTeamsCreation: boolean;
    UrlPrefix: string;
    AddUserSetPrefix: boolean;
    UserPrefixFormTitle: string;
    SharePointTemplatesIds: number[];
    TeamsTemplatesIds: number[];
}