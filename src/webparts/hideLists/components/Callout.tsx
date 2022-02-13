import * as React from 'react';
import { mergeStyleSets, Callout, DirectionalHint, MessageBar, MessageBarButton } from 'office-ui-fabric-react';

export interface ICalloutProps{
    _onCalloutDismiss: any;
    isCalloutVisible: any;
    target: string;
    className?: string;
}
export interface ICalloutState{

}
const styles = mergeStyleSets({
    callout: {
        maxWidth: 300,
    }
});

export default class CalloutComponent extends React.Component<ICalloutProps, ICalloutState>{
    constructor(props: ICalloutProps){
        super(props);
        this.state = {

        };
    }
    public render() {
        const { isCalloutVisible, target, children, className } = this.props;
        return(
            <Callout target={target} 
                className={`${styles.callout} ${className}`}
                isBeakVisible={false} onDismiss={this.props._onCalloutDismiss}
                hidden={!(isCalloutVisible)} coverTarget={true}
                directionalHint={DirectionalHint.topCenter} preventDismissOnScroll={true}>
                </Callout>
        );
    }
}