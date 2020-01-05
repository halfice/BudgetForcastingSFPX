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
import * as Datetime from 'react-datetime';
import 'react-datetime/css/react-datetime.css';
import styles from './Arabicformwebpart.module.scss';
import { default as pnp, ItemAddResult, Web } from "sp-pnp-js";
import ReactFileReader from 'react-file-reader';
import * as jquery from 'jquery';
import NumberFormat from 'react-number-format';
const logoj: any = require('./adda.jpg');


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
      BudgetForcasting: [],
      Remarks: "",


    }
    this._onChange = this._onChange.bind(this);
    this.OnchangeRemarks = this.OnchangeRemarks.bind(this);
    this.AddingProject = this.AddingProject.bind(this);
    this.handleInputChangeProjectName = this.handleInputChangeProjectName.bind(this);
    this.handleInputChangeForcastAmount = this.handleInputChangeForcastAmount.bind(this);
    this.AddActivity = this.AddActivity.bind(this);
  };
  OnchangeRemarks(event: any): void {
    this.setState({ Remarks: event.target.value });
  }


  AddActivity() {
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
      // console.log("Succes");
    }).catch(function (data) {
    });
    //AddingActivities End
  }

  handleInputChangeProjectName(event: any): void {
    this.setState({
      ProjectName: event.target.value
    });
  }

  handleInputChangeForcastAmount(event: any): void {
    this.setState({
      AmountForcast: event.target.value
    });
  }


  GenerateGuid() {
    var date = new Date();
    var guid = date.valueOf();
    return guid;
  }

  componentDidMount() {
    //console.log(this.state.pageContext.cultureInfo.currentCultureName);
    if (this.state.pageContext.cultureInfo.currentCultureName == "ar-SA") {
      this.setState({ LanguageKey: true });
    } else {
      this.setState({ LanguageKey: false });
    }
    this.GetUSerDetails();
    this.fetchProjects();
  }

  fetchProjects() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    var TempComplteDropDown = [];
    webx.lists.getByTitle("Projects").items.filter("Department eq '" + this.state.Department + "'").get().then((items: any[]) => {
      if (items.length > 0) {
        for (var i = 0; i < items.length; i++) {

          var NewData = {
            TotalAMount: items[i].AmountForCast,
            Title: items[i].Title,
          }
          if (i == 0) {
            var NewData1 = {
              TotalAMount: "0",
              Title: "Select Project",
            }
            TempComplteDropDown.push(NewData1);
          }
          TempComplteDropDown.push(NewData);
        }
        this.setState({
          ProjectsArray: TempComplteDropDown
        });
      }
    });
  }

  AddingProject() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    webx.lists.getByTitle("Projects").items.add({
      Title: this.state.ProjectName,
      AmountForCast: this.state.AmountForcast,
      Department: this.state.Department,
    }).then((iar: ItemAddResult) => {
      this.fetchProjects();
    });
  }

  AddForcastMonth() {
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

  FetchForCasting(ParamProjectName) {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    var TempComplteDropDown = [];
    webx.lists.getByTitle("Forcasting").items.select('AmountMonthly,ID,Project,Amount,Month,AmountMonthly,Department,Remaining,Delivered')
      .filter("Department eq '" + this.state.Department + "' and Project eq '" + ParamProjectName + "'").get().then((items: any[]) => {
        if (items.length > 0) {
          for (var i = 0; i < items.length; i++) {
            var NewData = {
              TotalAMount: items[i].Amount,
              Title: items[i].Title,
              Project: items[i].Project,
              Amount: items[i].Amount,
              Month: items[i].Month,
              AmountMonthly: items[i].AmountMonthly,
              Department: items[i].Department,
              Remaining: (parseFloat(items[i].Remaining)).toString(),
              Delivered: items[i].Delivered,
              ItemId: items[i].Id,
            }
            TempComplteDropDown.push(NewData);
          }
          this.setState({
            BudgetForcasting: TempComplteDropDown
          });
        } else {
          this.setState({
            BudgetForcasting: []
          });
        }

      });
  }

  handleFiles = files => {
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
        .then(function (data) {
          var RelativeUrls = "MyDocs/" + FinalName;//files.fileList[0].name;
          webx.getFolderByServerRelativeUrl(RelativeUrls).getItem().then(item => {
            // updating Start
            TemFileGuidName[0] = files.fileList[0].name + "|" + item["ID"];
            webx.lists.getByTitle("MyDocs").items.getById(item["ID"]).update({
              Guid: guid.toString(),
              ActualName: files.fileList[0].name
            }).then(r => {
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
    var TmpValue = this.state.LanguageKey ? false : true
    this.setState({
      LanguageKey: TmpValue,
    });
  }

  handleMain() {
    this.setState(
      {
        Screen: "Main",
      });
  }

  handleforcast() {
    this.setState(
      {
        Screen: "Forcast",
      });
  }

  handledeliverables() {
    this.setState(
      {
        Screen: "Deliverables",
      });
  }

  handleActivities() {
    this.setState(
      {
        Screen: "Activities",
      });
  }

  handleAddProject() {
    this.setState(
      {
        Screen: "AddProject",
      });
  }

  handleReport() {
    this.setState(
      {
        Screen: "Reports",
      });
  }

  onChangeProjectDropDown(event: any): void {

    var tmp = event.target.value;
    var TempArray = this.state.ProjectsArray;
    TempArray = TempArray.filter(function (TempArray) {
      return TempArray["Title"] == tmp;
    });
    var CurrentReportStatus = TempArray[0]["TotalAMount"];
    this.setState(
      {
        ProjectName: tmp,
        TotalAmountForcasted: CurrentReportStatus,
        BudgetForcasting: [],
      });
    this.FetchForCasting(tmp);
  }

  onChangeMonthDropDown(event: any): void {
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
    var SubProjectArrays = this.state.ProjectsArray.map(function (item, i) {
      return <option value={item["Title"]} key={item["Id"]}>{item["Title"]}</option>
    });

    var months = new Array("Select Month", "January", "February", "March",
      "April", "May", "June", "July", "August", "September",
      "October", "November", "December");
    var MonthsArray = months.map(function (item, i) {
      return <option value={item} key={item}>{item}</option>
    });

    var SubProjectArraysCards = this.state.BudgetForcasting.map(function (item, i) {
      return (

        <div className="col-md-4">
          <div className="card" >
            <img className="card-img-top" src={logoj} alt="Card image cap" />
            <div className="card-body">
              <h5 className="card-title">{strings.ProjectName}:{item["Project"]}</h5>
              <p className="card-text"> <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '}
                allowLeadingZeros={false} value={item["Amount"]}
              /></p>
              <p className="card-text"><NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '}
                value={item["Remaining"]}
                allowLeadingZeros={false} /></p>

              <a href="#" className="btn btn-primary">{item["Month"]}</a>
              <a href="#" className="btn btn-success">{item["AmountMonthly"]}</a>

            </div>
          </div>



        </div>



      );
    });




    return (

      <div className={this.state.LanguageKey == true ? styles.containerar : styles.containeren}>
        <div className={styles.mainHeading}>
          <div className={styles.mainHeading} onClick={this.handleMain.bind(this)}> {strings.greetings}</div>
          <div className={styles.Profile}>{strings.Deapartmentstr} : {this.state.Department}</div>
        </div>
        {
          this.state.Screen == "Main" &&
          <div className={styles.MainDiv} >
            <div className={styles.VideoMainDiv} onClick={this.handleforcast.bind(this)}>
              <span className={styles.innerspan}>Forcast</span>
            </div>
            <div className={styles.PhotosDiv} onClick={this.handledeliverables.bind(this)}>
              <span className={styles.innerspan} >Add Deliverables</span>
            </div>
            <div className={styles.VideoMainDivSearch} onClick={this.handleActivities.bind(this)}>
              <span className={styles.innerspan}>Activities</span>
            </div>
            {this.state.IsAuditorIsAdmin == true &&
              <div className={styles.VideoMainDivKPI} onClick={this.handleReport.bind(this)}>
                <span className={styles.innerspan}>Report (KPI) - Charts</span>
              </div>
            }
            {this.state.IsAuditorIsAdmin == true &&
              <div className={styles.VideoMainDivDelegation} onClick={this.handleAddProject.bind(this)}>
                <span className={styles.innerspan}>Add Projet</span>
              </div>
            }


          </div>
        }

        {
          this.state.Screen == "RequestProject" &&
          <div>
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
          <div >
            <div className={styles.PaddingForBottom}>
              <div>{strings.AdingProject}</div>
              <div className={styles.labelc}>{strings.ProjectName}</div>
              <input type="text" className={styles.textClass} id="txtPropjectName" onChange={this.handleInputChangeProjectName} />
              <div className={styles.labelc}>{strings.Amountforcast}</div>
              <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '} onValueChange={(values) => {
                var { formattedValue, value } = values;
                formattedValue = formattedValue.replace("aed", "");
                this.setState({ AmountForcast: formattedValue })
              }} />
            </div>
            <Stack horizontal >
              <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.AddingProject.bind(this)} />
            </Stack>
          </div>
        }

        {
          this.state.Screen == "Forcast" &&
          <div> <h3>Forcastng Amount :
           <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '}
              value={this.state.TotalAmountForcasted}
            />


          </h3>
            <div className={styles.PaddingForBottom}>
              <div>{strings.AdingProject}</div>
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
                this.setState({ MonthlyForcastAmount: formattedValue })
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
          <div> <h3>Forcastng Amount :
           <NumberFormat className={styles.textClass} thousandSeparator={true} prefix={'aed '}
              value={this.state.TotalAmountForcasted}
            />


          </h3>
            <div className={styles.PaddingForBottom}>
              <div>{strings.AdingProject}</div>
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
                this.setState({ MonthlyForcastAmount: formattedValue })
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
          this.state.Screen == "Activities" &&
          <div>
            <div className={styles.PaddingForBottom}>
              <div>{strings.AdingProject}</div>
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

              <div className={styles.labelc}>{strings.DescriptionFieldLabel}</div>
              <textarea value={this.state.Remarks} className={styles.myinputTextArea} onChange={this.OnchangeRemarks.bind(this)} >
                Hello there, this is some text in a text area
                        </textarea>
            </div>
            <Stack horizontal >
              <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.AddActivity.bind(this)} />
            </Stack>
          </div>
        }


        {
          this.state.Screen == "Reports" &&
          <div>
            <h1>Reports Amount</h1>
          </div>
        }

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
        var MBNumber = response.d.AccountName;
        var Departments = "IT";
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
}
