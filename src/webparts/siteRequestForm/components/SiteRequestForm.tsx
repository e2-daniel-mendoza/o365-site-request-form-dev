import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import styles from './SiteRequestForm.module.scss';
import { ISiteRequestFormProps } from './ISiteRequestFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { DefaultButton,TooltipHost,DirectionalHint } from 'office-ui-fabric-react';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';

import IRequestOptions from '../services/RequestOptions/IRequestOptions';
import RequestOptionsService from '../services/RequestOptions/RequestOptionsService';
import { ConsoleListener } from '@pnp/logging';

import SiteService from '../services/SiteCreationFlow/SiteService';

import * as pnp from '@pnp/sp';
import { IWeb } from '@pnp/sp/webs';
import { sp } from '@pnp/sp';
import ISiteOptions from '../services/SiteCreationFlow/ISiteOptions';
import { stringIsNullOrEmpty } from '@pnp/common';

import { IconButton } from "office-ui-fabric-react/lib/Button";



// const teamoptions: IDropdownOption[] = [
//   { key: 'yes', text: 'Yes' },
//   { key: 'no', text: 'No' },
// ];

const privatepublicoptions: IChoiceGroupOption[] = [
  { key: 'public', text: 'Public - Any Coliban Water staff member can view the site' },
  { key: 'private', text: 'Private - Only site owners and site members can view the site' },
];

// const stackTokens: IStackTokens = { childrenGap: 20 };

export interface ISiteRequestFormState {
  siteName: string;
  siteDescription: string;
  siteOwners: [];
  goToPeople: [];
  siteMembers: [];
  createTeam: boolean;
  isPublic: boolean;
  siteURLName: string;
  options: IRequestOptions[];
  selectedOption: IRequestOptions;
  userPrefix: string;
  currentUserEmail: string;
}

enum PeopleTypes {
  SiteOwner,
  GoToPeople,
  SiteMembers
}

export default class SiteRequestForm extends React.Component<ISiteRequestFormProps, ISiteRequestFormState> {

  public constructor(props: ISiteRequestFormProps) {
    super(props);
    this.state = {
      currentUserEmail: "",
      siteName: "",
      siteDescription: "",
      siteOwners: [],
      goToPeople: [],
      siteMembers: [],
      createTeam: false,
      isPublic: true,
      siteURLName: "",
      options: [],
      userPrefix: "",
      selectedOption: {
        Title: "",
        RestrictSiteCreation: false,
        RestrictSiteCreationToUserUids: [],
        RequiresApproval: false,
        ApproversUids: [],
        AllowMSTeamsCreation: true,
        UrlPrefix: "",
        AddUserSetPrefix: false,
        UserPrefixFormTitle: "",
        SharePointTemplatesIds: [],
        TeamsTemplatesIds: [],
      }
    };
    
    this._submitOnClick.bind(this);
  }

  private getRequestFormOptions(): Promise<IRequestOptions[]> {
    var requestOptionsService = new RequestOptionsService("Request Options", this.props.context);

    return requestOptionsService.GetRequestOptions();
  }

  public componentDidMount() {
    this._getCurrentUser();
    console.log("get site request options");
    this.getRequestFormOptions().then(o => {
      this.setState({options: o});
    });
  }

  public render(): React.ReactElement<ISiteRequestFormProps> {
    return (
      <div className={ styles.siteRequestForm }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p>{this.state.selectedOption.Title}</p>
              <h2>Site Request Form</h2>
              <p className={ styles.subTitle }>Please fill out the form below to get started.</p>
              <br />
              <TooltipHost
                content="This is the Workspace Type. Add more information..."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Workspace Type
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <ChoiceGroup
                options={this.state.options.map((i,index) => {return {key: index.toString(), text: i.Title}; })}
                onChange={(eventItem, option) => {this.setState({selectedOption: this.state.options[option.key]});}}
                 />

              <TooltipHost
                content="This is the name of the SharePoint site. The name of the SharePoint site must be unique compared to all the other sites in the organisation."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Site Name
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <TextField
                onBlur={​​​​​ eventitem => {
                  this.setState({siteName:eventitem.currentTarget.value});
                  this._getStatusNum();
                }}
                validateOnLoad={false}
                validateOnFocusOut={true}
                onGetErrorMessage={ eventitem => {
                  if(this.state.siteName == "") { 
                    return "This cannot be empty.";
                  } if(this.statusNum == 2) {
                    return "This site already exists.";
                  } if(this.statusNum == 1) {
                    return "This site is currently being provisioned.";
                  } if(this.statusNum == 3) {
                    return "An error occurred while provisioning the site.";
                  } else {
                    return "";
                  }
                }}
              /> 

              <TooltipHost
                content="This is the description of the SharePoint site explaining a brief summary of its purpose."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Description
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <TextField
                multiline rows={3}
              />
              {this.state.selectedOption.AddUserSetPrefix && 
                <TextField
                  label={this.state.selectedOption.UserPrefixFormTitle}
                  onBlur={eventitem=> this.setState({userPrefix: eventitem.currentTarget.value})}
                  />
              }
              <br />

              <TooltipHost
                content="List people who can have administrative rights in this SharePoint site."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Site Owners
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <PeoplePicker 
                defaultSelectedUsers={[this.state.currentUserEmail]}
                context={​​this.props.context}​​
                personSelectionLimit={​​999}​​
                groupName={​​""}​​
                onChange={​​ownersArr => this._getPeoplePickerItems(ownersArr, PeopleTypes.SiteOwner)}​​
                showHiddenInUI={​​false}​​
                principalTypes={​​[PrincipalType.User]}​​
                resolveDelay={​​1000}
                onGetErrorMessage={ ownersArr => {
                  if(this.state.siteOwners == []) {
                    return "This field cannot be empty";
                  } else {
                    return "";
                  }
                }}​
              />

              <TooltipHost
                content="List Go To People. Add more information..."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Go To People
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <PeoplePicker 
                context={​​this.props.context}​​
                personSelectionLimit={999}​​
                groupName={​​""}​​
                onChange={​​gtpArr => this._getPeoplePickerItems(gtpArr, PeopleTypes.GoToPeople)}​​
                showHiddenInUI={​​false}​​
                principalTypes={​​[PrincipalType.User]}​​
                resolveDelay={​​1000}​
              />

              <TooltipHost
                content="List of people who can access this SharePoint site with basic capabilities."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Site Members
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <PeoplePicker
                context={​​this.props.context}​​
                personSelectionLimit={​​999}​​
                groupName={​​""}​​
                onChange={membersArr => this._getPeoplePickerItems(membersArr, PeopleTypes.SiteMembers)}​​
                showHiddenInUI={​​false}​​
                principalTypes={​​[PrincipalType.User]}​​
                resolveDelay={​​1000}​
              />
              <br />

              <TooltipHost
                content="You have the choice of creating a Microsoft Team associated with this SharePoint site."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Create Microsoft Team
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              {this.state.selectedOption.AllowMSTeamsCreation && 
                <Toggle 
                  onText="Create"
                  offText="Don't Create"
                />
              }

              <TooltipHost
                content="You have the choice of either making this SharePoint public or private to the rest of the organization."
                id="tooltip"
                directionalHint={DirectionalHint.rightCenter}
                calloutProps={{ gapSpace: 0 }}
                styles={{ root: { display: "inline-block" } }}
              >
              <h2 style={{ cursor: "pointer" }}>Public or Private
                <IconButton iconProps={{ iconName: 'InfoSolid' }} title="InfoSolid" ariaLabel="InfoSolid" />
              </h2>
              </TooltipHost>
              <ChoiceGroup
                defaultSelectedKey="B"
                options={privatepublicoptions}
              />
              <br />
              <DefaultButton 
                text="Submit"
                onClick={event => {this._submitOnClick();}}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _getPeoplePickerItems(items: any[], targetPeople: PeopleTypes) {
    var usernames: string[] = [];
    for (var i: number = 0; i < items.length; i++) {
      usernames.push(this._trimUsername(items[i].id));
    }
    ​​ console.log('Items:', usernames);

    if(targetPeople === PeopleTypes.GoToPeople) {
      console.log("found go to people");
    }
    else if(targetPeople === PeopleTypes.SiteMembers) {
      console.log("found members");
    }
    else if(targetPeople === PeopleTypes.SiteOwner) {
      console.log("found owners");
    }
    else {
      console.error("somethings cooked -> people type not recognized");
    }

  }

  private _trimUsername(usernameToTrim: string): string {
    var uarr: string[] = usernameToTrim.split("|");
    var output: string = uarr[uarr.length - 1];
    return output;
  }
  
  public statusNum: number;

  private async _getStatusNum() {
    var siteService: SiteService = new SiteService(this.props.context);
    var totalURL: string = this._GetCompleteURL();
    siteService.CheckIfSiteExists(totalURL).then(value => {return value;});
    this.statusNum = await siteService.CheckIfSiteExists(totalURL).then(value => { return value;});
  }

  private async _getCurrentUser() {
    this.setState({
      currentUserEmail: 
        await sp.web.currentUser
          .get()
          .then((user) => { 
            return user.Email;
          })
    });
  }

  private _submitOnClick(): void {
    var siteService: SiteService = new SiteService(this.props.context);

    var totalURL: string = this._GetCompleteURL();

    siteService.CheckIfSiteExists(totalURL).then(value => {return value;});

    var siteOptions: ISiteOptions = {
      SiteName: this.state.siteName,
      SiteDescription: this.state.siteDescription,
      SiteOwners: this.state.siteOwners,
      SiteMembers: this.state.siteMembers,
      CreateTeam: this.state.createTeam,
      IsPublic: this.state.isPublic,
      SiteURL: totalURL,
      SharePointTemplatesIds: this.state.selectedOption.SharePointTemplatesIds,
      TeamsTemplatesIds: this.state.selectedOption.TeamsTemplatesIds,
      Properties: []
    };

    siteService.Create(this.props.flowURL, siteOptions);
    
  }

  //TODO add thing to change whether url is teams or sites, pulled from selected site object, change string to use regex on XX.sharepoint.com
  private _getBaseURL(absoluteURL: string = this.props.context.pageContext.site.absoluteUrl): string {

    var addTeamsSuffix: boolean = this.state.createTeam;
    var addSitesSuffix: boolean = !addTeamsSuffix;
    var stringtofind: string = ".sharepoint.com/";
    var index: number = absoluteURL.search(stringtofind);
    var baseurl: string = absoluteURL.substring(0, (index + stringtofind.length));

    if(addSitesSuffix) {
      baseurl += "sites/";
    }
    else if (addTeamsSuffix) {
      baseurl += "teams/";
    }

    return baseurl;
  }

  private _GetCompleteURL(): string {
    var prefix = this.state.selectedOption.UrlPrefix;
    if(prefix[prefix.length - 1] === "-") {
      prefix += "-";
    }

    var uprefix = this.state.userPrefix;
    if(this.state.selectedOption.AddUserSetPrefix && uprefix !== "") {
      uprefix += "-";
      prefix += uprefix;
    }

    var intendedURL = this._getBaseURL() + prefix + this.state.siteName.replace(" ", "-");

    return intendedURL;
  }

}

