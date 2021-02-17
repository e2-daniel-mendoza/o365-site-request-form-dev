import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { values } from 'office-ui-fabric-react';

import IRequestOptions from '../RequestOptions/IRequestOptions';
import ISiteOptions from './ISiteOptions';

export default class SiteService {
    private _context: WebPartContext;

    constructor(webPartContext: WebPartContext) {
        this._context = webPartContext;
    }

    public statusNum: number;

    public Create(flowURL: string ,siteOptions: ISiteOptions): void {
        //TODO remove this for the property object for Flow URL
        //let flowURL="https://prod-08.australiasoutheast.logic.azure.com:443/workflows/c89623077cfd4290b860e064f5794260/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=pYS3LHK6K1IIPvZt6e_MypVFGPcR_FyM06nDmnZUOKo";
    
        const body: string = JSON.stringify(
            siteOptions
        );
    
        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Content-type', 'application/json');

        const httpClientOptions: IHttpClientOptions = {
            body: body,
            headers: requestHeaders
        };


        var test = this._context.httpClient.post(
            flowURL,
            HttpClient.configurations.v1,
            httpClientOptions)
            .then((response: HttpClientResponse): Promise<HttpClientResponse> => {
              console.log("Email sent.");
              return response.json();
            });
    }

    public async CheckIfSiteExists(url: string): Promise<number> {
        // api documentation: https://docs.microsoft.com/en-us/sharepoint/dev/apis/site-creation-rest#get-modern-site-status
        let siteManagerURL: string = `${this._context.pageContext.web.absoluteUrl}/_api/SPSiteManager/status?url='${encodeURIComponent(url)}'`;

        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Accept', 'application/json;odata=verbose');
        requestHeaders.append('odata-version', '');

        const httpClientOptions: IHttpClientOptions = {

            headers: requestHeaders,

        };

        var siteManagerRequest = this._context.httpClient.get(
            siteManagerURL,
            HttpClient.configurations.v1,
            httpClientOptions)
            .then((response: HttpClientResponse): Promise<any> => {
                return response.json();
            });

        this.statusNum = await siteManagerRequest.then(value => {

            return value.d.Status.SiteStatus;

        });
        
        return this.statusNum;
    }

}