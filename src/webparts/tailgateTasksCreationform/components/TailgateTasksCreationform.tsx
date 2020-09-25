import * as React from 'react';
import styles from './TailgateTasksCreationform.module.scss';
import { ITailgateTasksCreationformProps } from './ITailgateTasksCreationformProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { sp, ItemAddResult } from "@pnp/sp";
import pnp from "sp-pnp-js";
import { Web } from 'sp-pnp-js';
// import {Web} from "@pnp/sp/webs";
// import {Lists} from "@pnp/sp/lists";
// import {Items} from "@pnp/sp/items";
//import pnp from "sp-pnp-js";
//import "@pnp/sp/items";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

//import { IAttachmentFileInfo } from "@pnp/sp/attachments";
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { IBasePickerSuggestionsProps, NormalPeoplePicker, ValidationState } from 'office-ui-fabric-react/lib/Pickers';
export interface ITailgateTasksCreationformState {
  topicValue: string;
  descriptionValue: string;
  fileAttach: any;
  allpeoplePicker_User: any;
  allpeoplePicker2_User: any;
  approverUsers: any;
  SignoffUsers: any;
  errortopicValue: string;
  errordescriptionValue: string;
  errorfileAttach: any;
  currentUserId: any;
  // errorallpeoplePicker_User: any;
  errorapproverUsers: boolean;
  errorSignoffUsers: boolean;
}
export default class TailgateTasksCreationform extends React.Component<ITailgateTasksCreationformProps, ITailgateTasksCreationformState> {
  public curretSiteURL = new Web(this.props.context.pageContext.web.absoluteUrl);
  constructor(props: ITailgateTasksCreationformProps) {
    super(props);
    this.state = {
      topicValue: "",
      descriptionValue: "",
      fileAttach: [],
      allpeoplePicker_User: [],
      allpeoplePicker2_User: [],
      approverUsers: "",
      SignoffUsers: [],
      currentUserId: "",
      //Error 
      errortopicValue: "",
      errordescriptionValue: "",
      errorfileAttach: "",
      //  errorallpeoplePicker_User: "",
      errorapproverUsers: false,
      errorSignoffUsers: false
    }
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._getSignOffPeoplePickerItems = this._getSignOffPeoplePickerItems.bind(this);
    this.getCurrentUser();
    //this.fileUploadCallback = this.fileUploadCallback.bind(this);
  }
  public componentDidMount() {
    console.log("componentDidMount");
  }
  public componentWillMount() {
    console.log("componentWillMount");
    sp.setup({
      spfxContext: this.props.context
    });
  }
  private _getPeoplePickerItems(items: any[]) {
    this.setState({
      allpeoplePicker_User: items,
      errorapproverUsers: false,
    });
  }
  private _getSignOffPeoplePickerItems(items: any[]) {
    this.setState({ allpeoplePicker2_User: items, errorSignoffUsers: false });
  }

  private getCurrentUser() {
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      this.setState({
        currentUserId: UserId['Id']
      });
    });
  }
  private draftForm = (): void => {

    this.state.topicValue.trim().length > 0 ? "" : this.setState({ errortopicValue: "Topic is required" });
    this.state.descriptionValue.trim().length > 0 ? "" : this.setState({ errordescriptionValue: "Description is required" });
    this.state.fileAttach.length == 0 ? this.setState({ errorfileAttach: "Approvers is required" }) : "";
    // this.state.allpeoplePicker_User.length > 0 ? "" : this.setState({ errorapproverUsers: true });
    this.state.allpeoplePicker2_User.length > 0 ? "" : this.setState({ errorSignoffUsers: true });

    if (this.state.topicValue.trim().length > 0 && this.state.descriptionValue.trim().length > 0 && this.state.fileAttach && this.state.allpeoplePicker2_User.length > 0) {
      //var today=today.getDay()+"/"+(today.getMonth()+1)+"/"+today.getFullYear();
      let today = new Date().toISOString().slice(0, 10)
      // var nwUser = {
      //   Id: this.state.allpeoplePicker_User[0]["id"],
      //   Email: this.state.allpeoplePicker_User[0]["secondaryText"]
      // };
      sp.web.lists.getByTitle("TailgateTasksActivity").items.add({
        Topic: this.state.topicValue,
        ProcessType0: "Tailgate",
        TaskIdentifier: "Tailgate Topic",
        Description: this.state.descriptionValue,
        Status: "Draft",
        RequestDate: today,
        RequesterNameId: {
          results: [this.state.currentUserId]// User/Groups ids as an array of numbers
        },
        ApproversId: {
          results: [this.state.allpeoplePicker_User.length > 0 ?
            this.state.allpeoplePicker_User[0]["id"] : []]  // User/Groups ids as an array of numbers
        },
        SignoffsId: {
          results: [this.state.allpeoplePicker2_User[0]["id"]]  // User/Groups ids as an array of numbers
        },
        //   SignoffsId: this.state.allpeoplePicker2_User[0]["id"]
      })
        .then((disID: ItemAddResult) => {
          //   console.log("Add Items to List SuccessFully");     
          //   console.log(disID.data.Id);
          //   let item = sp.web.lists
          //   .getByTitle("TailgateTasksActivity")
          //   .items.getById(disID.data.Id)
          //   item.attachmentFiles.add("Test", this.state.fileAttach).then(result => {   
          //   console.log("File uploaded successfully...")   
          // }); 
          let item = sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(disID.data.Id);
          item.attachmentFiles.add(this.state.fileAttach.fileName, this.state.fileAttach).then(v => {
            console.log("File upload successfully...!");
            alert("Save as Draft Successfully..!");
            this.setState({
              topicValue: "",
              descriptionValue: "",
              fileAttach: [],
              allpeoplePicker_User: [],
              allpeoplePicker2_User: []
            })
          });
          // return new Promise((resolve, reject) => {
          //  // this.getDiscussionId().then((_discussionId: number) => {
          //     let item = pnp.sp.web.lists
          //       .getByTitle('TailgateTasksActivity')
          //       .items.getById(disID.data.Id);
          //     item.attachmentFiles.add("kkk",this.state.fileAttach).then(result => {
          //     //  LogService.logInfo(fileName, "File upload Successfully");
          //       var _validFileExtensions = [".jpg", ".jpeg", ".bmp", ".gif", ".png"];

          //       console.log("kkoooko");

          //     });
          //  // });
          // });   
        });
    }
  }

  fileUploadCallback = event => {
    const file = event.target.files[0];
    this.setState({
      fileAttach: file,
      errorfileAttach: ""
    });
  };

  private submitForm = (): void => {
    this.state.topicValue.trim().length > 0 ? "" : this.setState({ errortopicValue: "Topic is required" });
    this.state.descriptionValue.trim().length > 0 ? "" : this.setState({ errordescriptionValue: "Description is required" });
    this.state.fileAttach.length == 0 ? this.setState({ errorfileAttach: "Approvers is required" }) : "";
    this.state.allpeoplePicker_User.length > 0 ? "" : this.setState({ errorapproverUsers: true });
    this.state.allpeoplePicker2_User.length > 0 ? "" : this.setState({ errorSignoffUsers: true });

    if (this.state.topicValue.trim().length > 0 && this.state.descriptionValue.trim().length > 0 && this.state.fileAttach && this.state.allpeoplePicker2_User.length > 0) {
      //var today=today.getDay()+"/"+(today.getMonth()+1)+"/"+today.getFullYear();
      let today = new Date().toISOString().slice(0, 10);
      sp.web.lists.getByTitle("TailgateTasksActivity").items.add({
        Topic: this.state.topicValue,
        Description: this.state.descriptionValue,
        Status: "Submit",
        ProcessType0: "Tailgate",
        TaskIdentifier: "Tailgate Topic",
        RequestDate: today,
        RequesterNameId: {
          results: [this.state.currentUserId]// User/Groups ids as an array of numbers
        },
        ApproversId: {
          results: [this.state.allpeoplePicker_User.length > 0 ? this.state.allpeoplePicker_User[0]["id"] : []]  // User/Groups ids as an array of numbers
        },

        SignoffsId: {
          results: [this.state.allpeoplePicker2_User[0]["id"]]  // User/Groups ids as an array of numbers
        },
      })
        .then((disID: ItemAddResult) => {
          //   console.log("Add Items to List SuccessFully");     
          //   console.log(disID.data.Id);
          //   let item = sp.web.lists
          //   .getByTitle("TailgateTasksActivity")
          //   .items.getById(disID.data.Id)
          //   item.attachmentFiles.add("Test", this.state.fileAttach).then(result => {   
          //   console.log("File uploaded successfully...")   
          // }); 
          let item = sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(disID.data.Id);
          item.attachmentFiles.add(this.state.fileAttach.fileName, this.state.fileAttach).then(v => {
            console.log("File upload successfully...!");
            alert("Submitted Successfully..!");
            this.setState({
              topicValue: "",
              descriptionValue: "",
              fileAttach: [""],
              allpeoplePicker_User: [""],
              allpeoplePicker2_User: [""]
            })
          });
          // return new Promise((resolve, reject) => {
          //  // this.getDiscussionId().then((_discussionId: number) => {
          //     let item = pnp.sp.web.lists
          //       .getByTitle('TailgateTasksActivity')
          //       .items.getById(disID.data.Id);
          //     item.attachmentFiles.add("kkk",this.state.fileAttach).then(result => {
          //     //  LogService.logInfo(fileName, "File upload Successfully");
          //       var _validFileExtensions = [".jpg", ".jpeg", ".bmp", ".gif", ".png"];

          //       console.log("kkoooko");

          //     });
          //  // });
          // });   
        });
    }

  }
  public render(): React.ReactElement<ITailgateTasksCreationformProps> {
    return (
      <div className={styles.tailgateTasksCreationform}>
        <h1>Tailgate JBC - New</h1>
        <div className={styles.container}>
          <hr></hr>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <TextField label="Topic" required
                value={this.state.topicValue}
                onChanged={newVal => {
                  newVal && newVal.length > 0
                    ? this.setState({
                      topicValue: newVal,
                      errortopicValue: ""
                    })
                    : this.setState({
                      topicValue: newVal,
                      errortopicValue:
                        "Topic is required"
                    });
                }}
                errorMessage={this.state.errortopicValue}
              />
            </div>
            <div className={styles.col_6}>
              <Label required>Approvers</Label>
              <PeoplePicker
                context={this.props.context}
                titleText=""
                personSelectionLimit={1}
                groupName={""}
                showtooltip={false}
                // isRequired={true}
                disabled={false}
                ensureUser={true}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}

              />
              {/* {this.state.errorapproverUsers ? <Label className={styles.pickerlabelErrormsg}>Approvers is required</Label> : ""} */}
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <TextField label="Description" required
                value={this.state.descriptionValue}
                onChanged={newDesVal => {
                  newDesVal && newDesVal.length > 0
                    ? this.setState({
                      descriptionValue: newDesVal,
                      errordescriptionValue: ""
                    })
                    : this.setState({
                      descriptionValue: newDesVal,
                      errordescriptionValue:
                        "Description is required"
                    });
                }}
                multiline rows={3} errorMessage={this.state.errordescriptionValue} />
            </div>
            <div className={styles.col_6}>
              <Label required>Sign offs</Label>
              <PeoplePicker
                //  peoplePickerCntrlclassName={styles.pickerErrormsg}
                context={this.props.context}
                titleText=""
                personSelectionLimit={1}
                groupName={""}
                showtooltip={false}
                //  isRequired={true}
                disabled={false}
                ensureUser={true}
                selectedItems={this._getSignOffPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              //errorMessage={this.state.SignoffUsers}
              />
              {this.state.errorSignoffUsers ? <Label className={styles.pickerlabelErrormsg}>Sign Offs is required</Label> : ""}

            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <div>
                <Label required>Attachment</Label>
                <input type="file" multiple accept=".xlsx,.xls,.doc, image/*, .docx,.ppt, .pptx,.txt,.pdf" onChange={this.fileUploadCallback}
                />
                {this.state.errorfileAttach ? <Label className={styles.pickerlabelErrormsg}>Attachment is required</Label> : ""}
              </div>
              {/* <TextField type="file" label="Attachment" accept=".xlsx,.xls,.doc, .docx,.ppt, .pptx,.txt,.pdf" onChanged={this.fileUploadCallback} onChange={this.fileUploadCallback}
                value={this.state.fileAttach} required errorMessage={this.state.errorfileAttach} /> */}
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_3}>
              <PrimaryButton className={styles.btnDraft} text="Save as Draft" onClick={this.draftForm} />

            </div>
            <div className={styles.col_3}>

              <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitForm} />
            </div>
          </div>

        </div>
      </div>
    );
  }
}
