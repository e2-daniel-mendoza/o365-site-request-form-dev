export default interface ISiteOptions {
    SiteName: string;
    SiteDescription: string;
    SiteOwners: string[];
    SiteMembers: string[];
    CreateTeam: boolean;
    IsPublic: boolean;
    SiteURL: string;
    SharePointTemplatesIds: number[];
    TeamsTemplatesIds: number[];
    Properties: string[];
}