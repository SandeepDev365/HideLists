import { IListInfo } from '@pnp/sp/lists';

export interface IHideListsState {
    data: IListInfo[];
    rowData: any;
    loading: boolean;
    loadingText: string;
    isCalloutVisible: boolean;
    isConfirmCalloutMessage: string;
    isConfirmCallOutVisible: boolean;
}
