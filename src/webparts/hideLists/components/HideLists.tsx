import * as React from 'react';
import styles from './HideLists.module.scss';
import { IHideListsProps } from './IHideListsProps';
import { IHideListsState } from './IHideListsState';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration, ISPHttpClientOptions } from '@microsoft/sp-http';
import ReactTable from "react-table-6";
import 'react-table-6/react-table.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import { DefaultButton, FocusTrapCallout } from '@microsoft/office-ui-fabric-react-bundle';
import { CallOutMessages, buttonTexts } from "../../../Helper/Constant";
import { FocusZone, FocusZoneTabbableElements, MessageBarType, PrimaryButton, Stack } from 'office-ui-fabric-react';
import { MessageBar, MessageBarButton } from 'office-ui-fabric-react';

export default class HideLists extends React.Component<IHideListsProps, IHideListsState> {
  private columns = [
    {
      Header: "List Name",
      accessor: "Title",
      headerStyle: { whiteSpace: 'nowrap' },
      style: { whiteSpace: 'normal' },
      sortable: true,
      filterable: true,
      Cell: row => (
        <div className='trucateData'>{row.original.Title}</div>
      )
    },
    {
      Header: "List GUID",
      accessor: "Id",
      headerStyle: { whiteSpace: 'nowrap' },
      style: { whiteSpace: 'normal' },
      sortable: true,
      filterable: true,
      Cell: row => (
        <div className='trucateData'>{row.original.Id}</div>
      )
    },
    {
      Header: "Is Hidden",
      accessor: "Hidden",
      headerStyle: { whiteSpace: 'nowrap' },
      style: { whiteSpace: 'normal' },
      sortable: true,
      filterable: true,
      width: 100,
      Cell: row => (
        <div className='trucateData'>{row.original.Hidden.toString()}</div>
      )
    },
    {
      Header: "Action",
      accessor: "",
      headerStyle: { whiteSpace: 'nowrap' },
      style: { whiteSpace: 'normal', color: '#0460A9', fontWeight: 'bold', flexWrap: 'wrap' },
      width: 140,
      Cell: row => (
        <>
          {
            row.original.Hidden ?
              <DefaultButton id={"btn" + row.index} onClick={() => this.actionbtnClicked(row, buttonTexts.Unhide)}>{buttonTexts.Unhide}</DefaultButton>
              :
              <DefaultButton id={"btn" + row.index} onClick={() => this.actionbtnClicked(row, buttonTexts.hide)}>{buttonTexts.hide}</DefaultButton>
          }
        </>
      )
    }
  ];
  private tableInstance;

  constructor(props) {
    super(props);

    this.state = {
      data: [],
      rowData: null,
      loading: false,
      loadingText: "",
      isCalloutVisible: false,
      isConfirmCalloutMessage: "",
      isConfirmCallOutVisible: false
    };
  }

  public async componentDidMount() {
    //this.hideUnhideLoader(true);
    await this.GetLists();
    //this.hideUnhideLoader(false);
  }

  private hideUnhideLoader(flag: boolean) {
    this.setState({ loading: flag });
  }

  private async unHideList() {
    let row = this.state.rowData;
    console.log("In UnHideList method");
    console.log("row", row);
    console.log("List Title - ", row.original.Title);
    console.log("List GUID - ", row.original.Id);

    try {
      await sp.web.lists.getById(row.original.Id).update({
        Hidden: false
      });
      await this.GetLists();
    }
    catch (ex) {
      console.log('Error', ex);
    }
  }

  private async hideList() {
    let row = this.state.rowData;
    console.log("In HideList method");
    console.log("row", row);
    console.log("List Title - ", row.original.Title);
    console.log("List GUID - ", row.original.Id);

    try {
      await sp.web.lists.getById(row.original.Id).update({
        Hidden: true
      });
      await this.GetLists();
    }
    catch (ex) {
      console.log('Error', ex);
    }
  }

  private async GetLists(): Promise<any> {
    return sp.web.lists.get().then((lsData) => { //filter("Hidden eq false and BaseType ne 1")
      console.log("Total number of lists are " + lsData.length);
      console.log("data", lsData);
      this.setState({ loading: false, data: lsData });
    });
  }

  private actionbtnClicked = (row, btnText) => {
    switch (btnText) {
      case buttonTexts.hide: this.setState({ isConfirmCallOutVisible: true, isConfirmCalloutMessage: CallOutMessages.hideList, rowData: row }); break;
      case buttonTexts.Unhide: this.setState({ isConfirmCallOutVisible: true, isConfirmCalloutMessage: CallOutMessages.unHideList, rowData: row }); break;
    }
  }

  private _onCalloutDismiss = () => {
    this.setState({ isCalloutVisible: false, isConfirmCalloutMessage: "", isConfirmCallOutVisible: false });
  }

  private onConfirmationMessageYesClicked = (event) => {
    switch (this.state.isConfirmCalloutMessage) {
      case CallOutMessages.hideList: this.hideList(); this._onCalloutDismiss(); break;
      case CallOutMessages.unHideList: this.unHideList(); this._onCalloutDismiss(); break;
    }
  }

  private onConfirmationMessageNoClicked = (event) => {
    this._onCalloutDismiss();
  }

  public render(): React.ReactElement {
    let { loading, loadingText, isCalloutVisible, isConfirmCallOutVisible, isConfirmCalloutMessage, data } = this.state;
    let btnId = this.state.rowData ? "btn" + this.state.rowData.index : "";
    console.log("columns", this.columns);
    console.log("data", data);
    return (
      <div>
        {/* <Loader loading={loading} loadingText={loadingText}/> */}
        <div>
          Site Url: <b>{this.props.ctx.pageContext.web.absoluteUrl}</b><br />
          Total Number of lists in the Site are <b>{data.length}</b>
        </div>
        <br />
        <ReactTable
          columns={this.columns}
          data={data}
          minRows={0}
          defaultPageSize={5}
          pageSizeOptions={[5, 10, 15]}
          noDataText={"Sorry, No data to display!!!"}
        />

        {/* <CalloutComponent _onCalloutDismiss={this._onCalloutDismiss} isCalloutVisible={isConfirmCallOutVisible} target={'#calloutdiv'} className='displayFormCallout'>
          <MessageBar messageBarType={MessageBarType.warning} className='saveChanges' isMultiline={true} actions={
            <div className='text-right mt20'>
              <MessageBarButton className='button custButton' onClick={(event) => this.onConfirmationMessageYesClicked(event)}>Yes</MessageBarButton>
              <MessageBarButton className='button custButton' onClick={(event) => this.onConfirmationMessageNoClicked(event)}>No</MessageBarButton>
            </div>
          }>
            {isConfirmCalloutMessage}
          </MessageBar>
        </CalloutComponent> */}
        {isConfirmCallOutVisible && (
          <FocusTrapCallout
            className='ms-CalloutExample-callout'
            ariaLabelledBy={'callout-label-1'}
            ariaDescribedBy={'callout-description-1'}
            role={'alertdialog'}
            gapSpace={0}
            target={`#${btnId}`}
            onDismiss={this._onCalloutDismiss}
            setInitialFocus={true}
          >
            <MessageBar messageBarType={MessageBarType.warning} className='saveChanges' isMultiline={true} actions={
              // <div className='text-right mt20'>
              //   <MessageBarButton className='button custButton' onClick={(event) => this.onConfirmationMessageYesClicked(event)}>Yes</MessageBarButton>
              //   <MessageBarButton className='button custButton' onClick={(event) => this.onConfirmationMessageNoClicked(event)}>No</MessageBarButton>
              // </div>
              <FocusZone handleTabKey={FocusZoneTabbableElements.all} isCircularNavigation>
                <Stack className='button custButton' gap={8} horizontal>
                  <PrimaryButton onClick={(event) => this.onConfirmationMessageYesClicked(event)}>Yes</PrimaryButton>
                  <DefaultButton onClick={(event) => this.onConfirmationMessageNoClicked(event)}>No</DefaultButton>
                </Stack>
              </FocusZone>
            }>
              {isConfirmCalloutMessage}
            </MessageBar>
          </FocusTrapCallout>
        )}
      </div>
    );
  }
}
