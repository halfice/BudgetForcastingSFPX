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


    }
    this._onChange = this._onChange.bind(this);
    this.AddingProject = this.AddingProject.bind(this);


  };


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

  }

  AddingProject() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    console.log(NewSiteUrl);
    let webx = new Web(NewSiteUrl);
    webx.lists.getByTitle("Projects").items.add({
      Title: this.state.ProjectName,
      AmountForCast:this.state.AmountForcast,
   }).then((iar: ItemAddResult) => {
      console.log(iar);
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
    console.log(NewSiteUrl);
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

  public render(): React.ReactElement<IArabicformwebpartProps> {
    //this.context.pageContext
    // it is only available on render

    // <Toggle defaultChecked onText="Arabic" offText="English" onChange={this._onChange.bind(this)} />
    return (

      <div className={this.state.LanguageKey == true ? styles.containerar : styles.containeren}>
        <div className={styles.mainHeading}>
          <div className={styles.mainHeading}> {strings.greetings}</div>
        </div>
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
        <br>
        </br>

        <div>
          <div>{strings.AdingProject}</div>
          <div className={styles.labelc}>{strings.ProjectName}</div>
          <input type="text" className={styles.textClass} id={this.state.ProjectName} />
          <div className={styles.labelc}>{strings.Amountforcast}</div>
          <input type="text" className={styles.textClass} id="amountforcast" />
          <Stack horizontal >
            <PrimaryButton text={strings.Submitbtn} allowDisabledFocus onClick={this.AddingProject.bind(this)} />
          </Stack>







        </div>



      </div>
    );
  }
}
