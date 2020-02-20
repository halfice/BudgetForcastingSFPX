import * as React from 'react';
import 'bootstrap/dist/css/bootstrap.css';

import { IArabicformwebpartProps } from './IArabicformwebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PageContext } from '@microsoft/sp-page-context';
import * as strings from 'ArabicformwebpartWebPartStrings';
import { Label, TextField, DefaultButton, PrimaryButton, Stack, IStackTokens } from 'office-ui-fabric-react';

import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import {
  DetailsListLayoutMode, Link, MarqueeSelection, DetailsList, Selection, Image, ImageFit,
  SelectionMode, Spinner, SpinnerSize, Fabric, ColumnActionsMode, IColumn, CheckboxVisibility,
  Callout, Panel, PanelType, IContextualMenuItem, autobind, ContextualMenu, IContextualMenuProps, DirectionalHint,
  css
} from 'office-ui-fabric-react';


import * as Datetime from 'react-datetime';
import 'react-datetime/css/react-datetime.css';
import styles from './Arabicformwebpart.module.scss';
import { default as pnp, ItemAddResult, Web, ConsoleListener } from "sp-pnp-js";
import ReactFileReader from 'react-file-reader';
import * as jquery from 'jquery';
import NumberFormat from 'react-number-format';
const logoj: any = require('./adda.jpg');
import 'bootstrap/dist/css/bootstrap.min.css';
import Row from 'react-bootstrap/Row';
import Col from 'react-bootstrap/Col';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
initializeIcons(undefined, { disableWarnings: true });

export default class Arabicformwebpart extends React.Component<IArabicformwebpartProps, {}> {
  public state: IArabicformwebpartProps;
  constructor(props, context) {
    super(props);
    this.state = {
      textFieldId: "sdf",
      LanguageKey: true,
      description: "",
      greetings: "",
      ListName: "List Form Name",
      Date: "",
      CurrentLanauge: "",
      spHttpClient: this.props.spHttpClient,
      pageContext: this.props.pageContext,
      siteurl: this.props.siteurl,
      ItemGuid: this.GenerateGuid().toString(),
      loading: false,
      UploadedFilesArray: [],
      ProjectName: "",
      AmountForcast: 0,
      Screen: "Main",
      IsAuditorIsAdmin: true,
      Department: "IT",
      ProjectsArray: [],
      SelectedMonth: "",
      TotalAmountForcasted: "",
      MonthlyForcastAmount: "",
      MonthlyDeliveredAmount: "",
      BudgetForcasting: [],
      Remarks: "",
      BalanceForcastTotal: 0,
      BalanceDeliverTotal: 0,
      Monitoritems: [],
      MonitorColumns: [],
      MonitorIndex: 0,
      showPanel: false,
      ProjectArrayGrid: [],
      PanelScreen: "Activities",
    };
    this._onChange = this._onChange.bind(this);
    this.OnchangeRemarks = this.OnchangeRemarks.bind(this);
    this.AddingProject = this.AddingProject.bind(this);
    this.handleInputChangeProjectName = this.handleInputChangeProjectName.bind(this);
    this.handleInputChangeForcastAmount = this.handleInputChangeForcastAmount.bind(this);
    this.AddActivity = this.AddActivity.bind(this);
    this.handleUpdateProject = this.handleUpdateProject.bind(this);
    this._onItemInvoked2 = this._onItemInvoked2.bind(this);
    this._onItemInvokedGetProjectDetail = this._onItemInvokedGetProjectDetail.bind(this);
    this.onChangeProjectDropDownrpt = this.onChangeProjectDropDownrpt.bind(this);

  }
  public OnchangeRemarks(event: any): void {
    this.setState({ Remarks: event.target.value });
  }

  private _onItemInvoked2(item: any): void {
    var CompleteItemArray = this.state.Monitoritems;
    let filteredarray = CompleteItemArray.filter(person => person["index"] == item["index"]);

    this.setState({
      showPanel: true,
      CurrentItemId: item.Id,
      MonitorIndex: parseInt(item["index"]),
      //ProjectName: filteredarray[0]["ProjectName"],
    });
  }

  private _onItemInvokedGetProjectDetail(item: any): void {
    //alert(item.Title);
    if (this.state.BudgetForcasting.length == 0) {
      this.FetchForCasting(item.Title);
    }
    this.setState({
      showPanel: true,
      PanelScreen: "Project"

    });
  }


  public AddActivity() {
    //Adding Activitites
    var dateFormat = require('dateformat');
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    webx.lists.getByTitle("Project Activities").items.add({
      Title: this.state.ProjectName, //Project Name
      Month: this.state.SelectedMonth,
      Department: this.state.Department,
      Activity: this.state.Remarks,
    }).then((response) => {
      this.LoadActivities(this.state.ProjectName.toString());
      // console.log("Succes");
    }).catch(error => {
    });
    //AddingActivities End
  }

  public LoadActivities(tmp) {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    var TempComplteDropDown = [];
    this.fillmonitorcolumns();
    if (tmp == "0") {
      webx.lists.getByTitle("Project Activities").items.get().then((items: any[]) => {
        if (items.length > 0) {
          for (var i = 0; i < items.length; i++) {
            var NewData = {
              Activity: items[i].Activity,
              Month: items[i].Month,
              Project: items[i].Title,
            };
            TempComplteDropDown.push(NewData);
          }
          this.setState({
            Monitoritems: TempComplteDropDown
          });
        }
      });
    }
    else {
      webx.lists.getByTitle("Project Activities").items.filter(`Department eq'${this.state.Department}'and Title eq'${tmp}'`).get().then((items: any[]) => {
        if (items.length > 0) {
          for (var i = 0; i < items.length; i++) {
            var NewData = {
              Activity: items[i].Activity,
              Month: items[i].Month,
              Project: items[i].Title,
            };
            TempComplteDropDown.push(NewData);
          }
          this.setState({
            Monitoritems: TempComplteDropDown
          });
        }
      });
    }
  }

  public handleInputChangeProjectName(event: any): void {
    this.setState({
      ProjectName: event.target.value
    });
  }

  public handleInputChangeForcastAmount(event: any): void {
    this.setState({
      AmountForcast: event.target.value
    });
  }


  public GenerateGuid() {
    var date = new Date();
    var guid = date.valueOf();
    return guid;
  }

  public componentDidMount() {
    console.log(this.state.pageContext.cultureInfo.currentCultureName);
    if (this.state.pageContext.cultureInfo.currentCultureName == "ar-SA") {
      this.setState({ LanguageKey: true });
    } else {
      this.setState({ LanguageKey: false });
    }
    this.GetUSerDetails();
    this.fetchProjects();
  }

  public fetchProjects() {

    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    var TempComplteDropDown = [];
    var TempProjectGrid = [];
    webx.lists.getByTitle("Projects").items.filter("Department eq '" + this.state.Department + "'").get().then((items: any[]) => {
      if (items.length > 0) {
        for (var i = 0; i < items.length; i++) {

          var NewData = {
            TotalAMount: items[i].AmountForCast,
            Title: items[i].Title,
          };
          if (i == 0) {
            var NewData1 = {
              TotalAMount: "0",
              Title: strings.SelectString,
            };
            TempComplteDropDown.push(NewData1);
          }
          TempComplteDropDown.push(NewData);
          TempProjectGrid.push(NewData);
        }
        this.setState({
          ProjectsArray: TempComplteDropDown,
          ProjectArrayGrid: TempProjectGrid,
        });
      }
    }).catch(err => {
      console.log(err);
    });
  }

  public AddingProject() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    if (this.state.ProjectName == "" && this.state.ProjectName == null) {
      alert("Enter Project NAme");
      return;
    }
    webx.lists.getByTitle("Projects").items.add({
      Title: this.state.ProjectName,
      AmountForCast: this.state.AmountForcast,
      Department: this.state.Department,
    }).then((iar: ItemAddResult) => {
      this.fetchProjects();
    });
  }

  public handleUpdateProject() {

    var tmp = this.state.SelectedMonth;
    var TempArray = this.state.BudgetForcasting;
    //  TempArray = TempArray.filter(function (TempArray) {
    // return TempArray["Month"] == tmp;
    // });
    let filteredarray = TempArray.filter(person => person["Month"] == tmp);
    if (filteredarray != null) {
      var ItemID = 0;
      ItemID = filteredarray[0]["ItemId"];
      var NewISiteUrl = this.props.siteurl;
      var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
      let webx = new Web(NewSiteUrl);
      webx.lists.getByTitle("Forcasting").items.getById(ItemID).update({
        Delivered: this.state.MonthlyDeliveredAmount,
      }).then(r => {
        this.FetchForCasting(this.state.ProjectName);
      });
    }
  }

  public AddForcastMonth() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    webx.lists.getByTitle("Forcasting").items.add({
      Title: "ADDA-Budget Forcasting",
      Project: this.state.ProjectName,
      Amount: this.state.TotalAmountForcasted,
      Month: this.state.SelectedMonth,
      AmountMonthly: this.state.MonthlyForcastAmount,
      Department: this.state.Department,
    }).then((iar: ItemAddResult) => {
      this.FetchForCasting(this.state.ProjectName.toString());
    });
  }

  public FetchForCasting(ParamProjectName) {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    var TempComplteDropDown = [];
    var tmpBalance = 0;
    var tmpBalanceDlvr = 0;
    webx.lists.getByTitle("Forcasting").items.select('AmountMonthly,ID,Project,Amount,Month,AmountMonthly,Department,Remaining,Delivered')
      .filter("Department eq '" + this.state.Department + "' and Project eq '" + ParamProjectName + "'").get().then((items: any[]) => {
        if (items.length > 0) {
          for (var i = 0; i < items.length; i++) {

            var TmpDevlier = 0;
            if (items[i].Delivered != "" && items[i].Delivered != null) {
              TmpDevlier = items[i].Delivered;
            }

            var NewData = {
              TotalAMount: items[i].Amount,
              Title: items[i].Title,
              Project: items[i].Project,
              Amount: items[i].Amount,
              Month: items[i].Month,
              AmountMonthly: items[i].AmountMonthly,
              Department: items[i].Department,
              Remaining: (parseFloat(items[i].Remaining)).toString(),
              Delivered: TmpDevlier,
              ItemId: items[i].Id,
            };
            TempComplteDropDown.push(NewData);
            var TempAmountMonthly = items[i].AmountMonthly;
            var tmpFloatAmount = parseFloat(TempAmountMonthly);
            var tmpTotalAMount = parseFloat(items[i].Amount);
            tmpBalance = tmpBalance + tmpFloatAmount;


            tmpBalanceDlvr = tmpBalanceDlvr + TmpDevlier;

          }



          tmpBalance = tmpTotalAMount - tmpBalance;
          tmpBalanceDlvr = tmpTotalAMount - tmpBalanceDlvr;


          this.setState({
            BudgetForcasting: TempComplteDropDown,
            BalanceForcastTotal: tmpBalance,
            BalanceDeliverTotal: tmpBalanceDlvr
          });
        } else {
          this.setState({
            BudgetForcasting: []
          });
        }

      });
  }

  public handleFiles = files => {
    var TemFileGuidName = [];
    var component = this;
    component.setState({ loading: true });
    var FileExtension = this.getFileExtension1(files.fileList[0].name);
    var date = new Date();
    var guid = date.valueOf();
    if (this.state.ItemGuid == "-1") {
      this.setState({ ItemGuid: guid });
    }
    //alert(this.state.ItemGuid);   
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);

    var FinalName = guid + FileExtension;
    var binary_string = window.atob(files.base64.split(',')[1]);
    var len = binary_string.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) {
      bytes[i] = binary_string.charCodeAt(i);
    }
    var myBlob = bytes.buffer;
    webx.get().then(r => {
      webx.getFolderByServerRelativeUrl("MyDocs")
        .files.add(FinalName.toString(), myBlob, true)
        .then(data => {
          var RelativeUrls = "MyDocs/" + FinalName;//files.fileList[0].name;
          webx.getFolderByServerRelativeUrl(RelativeUrls).getItem().then(item => {
            // updating Start
            TemFileGuidName[0] = files.fileList[0].name + "|" + item["ID"];
            webx.lists.getByTitle("MyDocs").items.getById(item["ID"]).update({
              Guid: guid.toString(),
              ActualName: files.fileList[0].name
            }).then(() => {
              component.setState({ loading: false });
              component.setState({ UploadedFilesArray: component.state.UploadedFilesArray.concat(TemFileGuidName) });
            });
          }); //Retrive Doc Info End


        });
    });
  }

  private getFileExtension1(filename) {
    return (/[.]/.exec(filename)) ? /[^.]+$/.exec(filename)[0] : undefined;
  }

  public onSelectDateRequired(event: any): void {
    this.setState({ RequiredDate: event._d });
  }

  private _onChange(item: any): void {
    var TmpValue = this.state.LanguageKey ? false : true;
    this.setState({
      LanguageKey: TmpValue,
    });
  }

  public handleMain() {
    this.setState(
      {
        Screen: "Main",
      });
  }

  public handleforcast() {
    this.setState(
      {
        Screen: "Forcast",
      });
  }

  public handledeliverables() {
    this.setState(
      {
        Screen: "Deliverables",
      });
  }

  public handleActivities() {
    this.setState(
      {
        Screen: "Activities",
      });
  }

  public handleAddProject() {
    this.setState(
      {
        Screen: "AddProject",
      });
    this.fillmonitorcolumns();
  }

  public handleReport() {
    this.LoadActivities("0");
    this.setState(
      {
        Screen: "Reports",
      });
  }


  public filterItems = (arr, query) => {
    return arr.filter(el => el.Title == query);
  }

  public onChangeProjectDropDown(event: any): void {
    var tmp = event.target.value;
    var TempArray = this.state.ProjectsArray;
    var newar = this.filterItems(TempArray, tmp);



    var CurrentReportStatus = newar[0]["TotalAMount"];
    this.setState(
      {
        ProjectName: tmp,
        TotalAmountForcasted: CurrentReportStatus,
        BudgetForcasting: [],
      });
    this.FetchForCasting(tmp);
  }

  public onChangeProjectDropDownrpt(event: any): void {
    var tmp = event.target.value;

    var TempArray = this.state.ProjectsArray;
    var newar = this.filterItems(TempArray, tmp);



    var CurrentReportStatus = newar[0]["TotalAMount"];
    this.setState(
      {
        ProjectName: tmp,
        TotalAmountForcasted: CurrentReportStatus,
        BudgetForcasting: [],
      });
    this.LoadActivities(tmp);
  }

  public onChangeMonthDropDown(event: any): void {
    var tmp = event.target.value;
    this.setState(
      {
        SelectedMonth: tmp,
      });
  }



  public render(): React.ReactElement<IArabicformwebpartProps> {
    //this.context.pageContext
    // it is only available on render
    var defaultValue = 'My default value';
    // <Toggle defaultChecked onText="Arabic" offText="English" onChange={this._onChange.bind(this)} />

    // <Toggle defaultChecked onText="Arabic" offText="English" onChange={this._onChange.bind(this)} />
    var SubProjectArrays = this.state.ProjectsArray.map((item, i) => {
      return <option value={item["Title"]} key={item["Id"]}>{item["Title"]}</option>;
    });




    var months = new Array("-", "January", "February", "March",
      "April", "May", "June", "July", "August", "September",
      "October", "November", "December");


    var MonthsArray = months.map((item, i) => {
      return <option value={item} key={item}>{item}</option>;
    }); // months.map(function (item, i) {


    //filling the Panel for Project start
    if (this.state.PanelScreen == "Project") {

      
      var Panelhtml = this.state.BudgetForcasting.map((item, i, arr) => {
        return (
          <div>
            <Row>
              <Col>

                {item["Month"]}
              </Col>
              <Col>
                
                 Forcast : {item["AmountMonthly"]}
                
              </Col>
              <Col>
                
                 Delivered : {item["Delivered"]}
                
              </Col>
            </Row>


          </div>
        );
      });
      
      var PanelFooter = this.state.BudgetForcasting.map((item, i) => {
        return (
          <div>
            <Row>
              <Col>Remaining :{item["Title"]}</Col>
              <Col>Total Amount :{item["Amount"]}</Col>
            </Row>

          </div>);

      });
    }
    //filling the panel for project end



    if (this.state.Screen == "Forcast") {
      var SubProjectArraysCards = this.state.BudgetForcasting.map((item, i) => {
        return (
          <div className={styles.circleContent}>
            <span >
              {item["Month"]}
            </span>
            <div >
              {item["AmountMonthly"]}
            </div>
          </div>
        );
      });
      //
    }

    if (this.state.Screen == "Deliverables") {
      var SubProjectArraysCardsDelivered = this.state.BudgetForcasting.map((item, i) => {
        return (
          <div className={styles.circleContentdeliver}>
            <span >
              M: {item["Month"]}
            </span>
            <span>
              F: {item["AmountMonthly"]}
            </span>
            <span >
              D: {item["Delivered"]}
            </span>
          </div>
        );
      });
      //
      //
    }


    return (

      <div className={this.state.LanguageKey == true ? styles.containerar : styles.containeren}>
        <div className={styles.MainDivClass} >
          <Row>
            <Col>
              <span className="glyphicon glyphicon-home"></span>
              <div className={styles.mainHeading} onClick={this.handleMain.bind(this)}> <Icon iconName='AddHome' /></div></Col>
            <Col><div className={styles.mainHeading}><Icon iconName='PartyLeader' /> {this.state.Department}</div></Col>
          </Row>
        </div>
        {
          this.state.Screen == "Main" &&
          <div className={styles.MainDiv}>
            <Row>
              <Col>

                <div className={styles.maindivAddProject} onClick={this.handleAddProject.bind(this)}>

                  <Icon iconName='AdminALogoInverse32' />
                  <hr></hr>
                  {strings.mainproject}</div>
              </Col>


              <Col>
                <div className={styles.maindivAddProject} onClick={this.handleforcast.bind(this)}>
                  <Icon iconName='Trending12' />
                  <hr></hr>
                  {strings.mainforcast}</div>
              </Col>


              <Col>
                <div className={styles.maindivAddProject} onClick={this.handledeliverables.bind(this)}>
                  <Icon iconName='ReleaseDefinition' />
                  <hr></hr>
                  {strings.maindeliverables}</div>
              </Col>





              <Col> <div className={styles.maindivAddProject} onClick={this.handleActivities.bind(this)}>
                <Icon iconName='TeamsLogoInverse' />
                <hr></hr>
                {strings.mainactivities}</div></Col>
              <Col> <div className={styles.maindivAddProject} onClick={this.handleReport.bind(this)}>
                <Icon iconName='ReportAdd' />
                <hr></hr>
                {strings.mainreport}</div></Col>

            </Row>







          </div>
        }

        {
          this.state.Screen == "RequestProject" &&
          <div className={styles.MainDivClass} >
            <h3> {strings.RequestType}</h3>
            <div className={styles.containerinner}>
              <div className={styles.labelc}>{strings.Title}</div>
              <input type="text" className={styles.textClass} id="myTextareaas" />
              <br></br>
              <div className={styles.labelc}>{strings.Title}</div>
              <input type="text" className={styles.textClass} id="myTextareaas" />
              <br></br>
              <div className={styles.rowDate}>
                <div className={styles.labelc}>{strings.Date}</div>
                <div >
                  <Datetime onChange={this.onSelectDateRequired.bind(this)} />
                </div>
              </div>

              <div className={styles.row}>
                <ReactFileReader fileTypes={[".csv", ".xlsx", ".Docx"]} handleFiles={this.handleFiles.bind(this)} base64={true} >
                  <button className='btn'>{strings.Upload}</button>
                </ReactFileReader>
              </div>
            </div>
            <hr></hr>
            <Stack horizontal >
              <DefaultButton text={strings.Submitbtn} allowDisabledFocus />
              <PrimaryButton text={strings.Cancelbtn} allowDisabledFocus />
            </Stack>

          </div>
        }

        {
          this.state.Screen == "AddProject" &&
          <div className={styles.MainDivClass} >
            <div className={styles.PaddingForBottom}>
              <div>{strings.AdingProject}</div>
              <div className={styles.labelc}>{strings.ProjectName}</div>
              <input type="text" className={styles.textClass} id="txtPropjectName" onChange={this.handleInputChangeProjectName} />
              <div className={styles.labelc}>{strings.Amountforcast}</div>
              <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '} onValueChange={(values) => {
                var { formattedValue, value } = values;
                formattedValue = formattedValue.replace("aed", "");
                this.setState({ AmountForcast: formattedValue });
              }} />
            </div>
            <Stack horizontal >
              <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.AddingProject.bind(this)} />
            </Stack>

            <Row>
              <Col>
                <h1>{strings.AdingProject}</h1>
                <hr>

                </hr>
                <div className={styles.containeren} >
                  <DetailsList
                    items={this.state.ProjectArrayGrid}
                    columns={this.state.MonitorColumns}
                    //  onRenderItemColumn={_renderItemColumnMonitor}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionPreservedOnEmptyClick={true}
                    ariaLabelForSelectionColumn="Toggle selection"
                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                    checkButtonAriaLabel="Row checkbox"



                    onItemInvoked={this._onItemInvokedGetProjectDetail}

                  />

                </div>
              </Col>


            </Row>
          </div>
        }

        {
          this.state.Screen == "Forcast" &&
          <div className={styles.MainDivClass} >
            <Row>
              <Col>
                <div className={styles.circle}> <p>Total</p>
                  {this.state.TotalAmountForcasted}
                </div>
              </Col>
              <Col>
                <div className={styles.circle}> <p>BAL</p>
                  {this.state.BalanceForcastTotal}
                </div>
              </Col>

            </Row>

            <div className={styles.PaddingForBottom}>

              <div className={styles.labelc}>{strings.ProjectName}</div>
              <select value={this.state.ProjectName} className={styles.myinputSelect}
                defaultValue={defaultValue}
                onChange={this.onChangeProjectDropDown.bind(this)}>{SubProjectArrays}
              </select>

              <div className={styles.labelc}>{strings.month}</div>
              <select value={this.state.SelectedMonth} className={styles.myinputSelect}
                defaultValue={defaultValue}
                onChange={this.onChangeMonthDropDown.bind(this)}>{MonthsArray}
              </select>

              <div className={styles.labelc}>{strings.Amountforcast}</div>
              <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '} onValueChange={(values) => {
                var { formattedValue, value } = values;
                formattedValue = formattedValue.replace("aed", "");
                this.setState({ MonthlyForcastAmount: formattedValue });
              }} />
            </div>
            <Stack horizontal >
              <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.AddForcastMonth.bind(this)} />
            </Stack>
            <hr>
            </hr>
            {
              this.state.BudgetForcasting.length > 0 &&
              <div className="row">
                {SubProjectArraysCards}
              </div>
            }

          </div>
        }



        {
          this.state.Screen == "Deliverables" &&
          <div className={styles.MainDivClass} >
            <Row>
              <Col>
                <div className={styles.circle}> <p>Total</p>
                  {this.state.TotalAmountForcasted}
                </div>
              </Col>
              <Col>
                <div className={styles.circle}> <p>BAL</p>
                  {this.state.BalanceForcastTotal}
                </div>
              </Col>
              <Col>
                <div className={styles.circle}> <p>Del</p>
                  {this.state.BalanceDeliverTotal}
                </div>
              </Col>

            </Row>



            <div className={styles.PaddingForBottom}>


              <div className={styles.labelc}>{strings.ProjectName}</div>
              <select value={this.state.ProjectName} className={styles.myinputSelect}
                defaultValue={defaultValue}
                onChange={this.onChangeProjectDropDown.bind(this)}>{SubProjectArrays}
              </select>


              <div className={styles.labelc}>{strings.month}</div>
              <select value={this.state.SelectedMonth} className={styles.myinputSelect}
                defaultValue={defaultValue}
                onChange={this.onChangeMonthDropDown.bind(this)}>{MonthsArray}
              </select>

              <div className={styles.labelc}>{strings.AmountDelivere}</div>
              <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '} onValueChange={(values) => {
                var { formattedValue, value } = values;
                formattedValue = formattedValue.replace("aed", "");
                this.setState({ MonthlyDeliveredAmount: formattedValue });
              }} />



            </div>




            <Stack horizontal >
              <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.handleUpdateProject.bind(this)} />
            </Stack>
            <hr>
            </hr>


            {
              this.state.BudgetForcasting.length > 0 &&
              <div className="row">
                {SubProjectArraysCardsDelivered}
              </div>
            }

          </div>
        }



        {
          this.state.Screen == "Activities" &&
          <div className={styles.MainDivClass} >
            <div className={styles.PaddingForBottom}>

              <div className={styles.labelc}>{strings.ProjectName}</div>
              <select value={this.state.ProjectName} className={styles.myinputSelect}
                defaultValue={defaultValue}
                onChange={this.onChangeProjectDropDownrpt.bind(this)}>{SubProjectArrays}
              </select>
              <div className={styles.labelc}>{strings.month}</div>
              <select value={this.state.SelectedMonth} className={styles.myinputSelect}
                defaultValue={defaultValue}
                onChange={this.onChangeMonthDropDown.bind(this)}>{MonthsArray}
              </select>

              <div className={styles.labelc}>{strings.DescriptionFieldLabel}</div>
              <textarea value={this.state.Remarks} className={styles.myinputTextArea} onChange={this.OnchangeRemarks.bind(this)} >
                Hello there, this is some text in a text area
                        </textarea>
            </div>
            <Stack horizontal >
              <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.AddActivity.bind(this)} />
            </Stack>

            <div>
              <DetailsList
                items={this.state.Monitoritems}
                columns={this.state.MonitorColumns}
                //  onRenderItemColumn={_renderItemColumnMonitor}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionPreservedOnEmptyClick={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                onItemInvoked={this._onItemInvoked2}

              />

            </div>
          </div>
        }


        {
          this.state.Screen == "Reports" &&
          <div className={styles.MainDivClass} >

            <DetailsList
              items={this.state.Monitoritems}
              columns={this.state.MonitorColumns}
              //  onRenderItemColumn={_renderItemColumnMonitor}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              onItemInvoked={this._onItemInvoked2}

            />
          </div>
        }


        <Panel
          isOpen={this.state.showPanel}
          type={PanelType.medium}
          onDismiss={this._onClosePanel}
          headerText="Details"
          closeButtonAriaLabel="Close"
        >
          <h1>Budget forcast </h1>

  {
    this.state.BudgetForcasting.length>0 &&
<Row>
<Col>Name :{this.state.BudgetForcasting[0]["Title"]}</Col>
              <Col>Total Amount :{this.state.BudgetForcasting[0]["Amount"]}</Col>
</Row>
  }
<Row>

{Panelhtml}
</Row>
          

        </Panel>

      </div>
    );
  }

  private GetUSerDetails() {
    var reactHandler = this;
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    var reqUrl = NewSiteUrl + "/_api/sp.userprofiles.peoplemanager/GetMyProperties";
    jquery.ajax(
      {
        url: reqUrl, type: "GET", headers:
        {
          "accept": "application/json;odata=verbose"
        }
      }).then((response) => {
        var Name = response.d.DisplayName;
        var email = response.d.Email;
        var oneUrl = response.d.PersonalUrl;
        var imgUrl = response.d.PictureUrl;
        var jobTitle = response.d.Title;
        var profUrl = response.d.UserUrl;
        var MBNumber = response.d.AccountName;
        var Departments = "IT";//"تكنولوجيا المعلومات";
        var Tmpe = MBNumber.toString().split('|');
        var Tmp2 = Tmpe[2].toString().split('@');
        MBNumber = Tmp2[0];
        reactHandler.setState({
          //EmployeeName: response.d.DisplayName,
          //EmployeeNumber: MBNumber,
          //EmployeeEmail: email,
          Department: Departments
        });

      });
  }

  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }


  public fillmonitorcolumns() {
    var Tempcolumns = [];
    if (this.state.Screen != "Main") {
      var ONMColus = ["Activity", "Month", "Project", "."];
      for (var i = 0; i < ONMColus.length; i++) {
        var newData = {
          key: ONMColus[i],
          name: ONMColus[i],
          fieldName: ONMColus[i],
          minWidth: 0,
          maxWidth: 0,
          isResizable: true,
          ariaLabel: 'activities',
          headerClassName: 'DetailsListExample-header--FileIcon',
        };

        Tempcolumns.push(newData);
      }
    } else {
      var xONMColus = ["Title", "TotalAMount"];
      for (var x = 0; x < xONMColus.length; x++) {
        var xnewData = {
          key: xONMColus[x],
          name: xONMColus[x],
          fieldName: xONMColus[x],
          ariaLabel: 'Projects',
          headerClassName: 'DetailsListExample-header--FileIcon',
        };

        Tempcolumns.push(xnewData);
      }
    }

    this.setState({
      MonitorColumns: Tempcolumns,
    });
  }


}