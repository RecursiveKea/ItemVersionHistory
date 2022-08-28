export interface IRestApiResponse {
    value: any;
    "@odata.nextLink"?: string;
    
}

export interface IRestApiResponseNext {
    status: number;
    ok: boolean;
    Response: IRestApiResponse;
}