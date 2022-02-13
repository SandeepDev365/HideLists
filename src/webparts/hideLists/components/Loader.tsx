import * as React from 'react';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface ILoaderprops {
    loading: boolean;
    loadingText: string;
}
export interface ILoaderState {
    loadingText: string;
}

export default class Loader extends React.Component<ILoaderprops, ILoaderState>
{
    public render() {
        return(
            <div>
                <Dialog hidden={this.props.loading}
                    dialogContentProps={{ type: DialogType.normal, title: '' }}
                    modalProps={{ isBlocking: true, containerClassName: 'ms-dialogMainOverride' }}
                >
                    <Spinner size={SpinnerSize.large} label={this.props.loadingText} ariaLive='assertive' />
                </Dialog>
            </div>
        );
    }
}