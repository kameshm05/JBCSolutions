import * as React from 'react';
import styles from './ReportSignOff.module.scss';
import { IReportSignOffProps } from './IReportSignOffProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import "@pnp/sp/attachments";
import "@pnp/sp/profiles";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { Checkbox, DefaultButton, Label, mergeStyles, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { IconButton, IIconProps, IContextualMenuProps, Stack, Link } from 'office-ui-fabric-react';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import {
  Dialog,
  DialogFooter,
  DialogType,
  IDialogStyles,
} from "office-ui-fabric-react/lib/Dialog";
import {IReportSignOffState } from './IReportSignOffState';
var signOffCheck=[];
import classnames from 'classnames';
import swal from 'sweetalert';
import '../../../ExternalRef/style.css';
export default class ReportSignOff extends React.Component<IReportSignOffProps,IReportSignOffState > {
  public state;
  public queryStringId: any;
  public queryMode:any;
  public currentUserEmail:any;
  public summaryStatus:any;
  public approverstextArray=[];
  public signOfftextArray=[];
  options: IChoiceGroupOption[];
  public dialogContentProps = {
    //subText: '',
    title: '',
  };
  constructor(props: IReportSignOffProps) {
    super(props);
   
    this.state = {
      topicValue: "",
      descriptionValue: "",
      allpeoplePicker_User: [],
      allpeoplePicker2_User: [],
      approverUsers: "",
      SignoffUsers: [],
      filePickerResult: [],
      //Error 
      errortopicValue: "",
      errordescriptionValue: "",
      errorfileAttach: "",
      errorapproverUsers: false,
      errorSignoffUsers: false,
      addformHide: false,
      getApprovepeoplePicker_User: [],
      getSignOffUser: [],
      queryStringId:0,
      headerContent:"Self Sign-Offs",
      viewMode:false,
      StatusSummary:"",
      fetchSignOffUsers:[],
      removedFiles:[],
      newFiles:[],
      fileDetails: [],
      approveStatus:"Approve",
      comments:"",
      errorcomments:"",
      fetchApprovers:[],
      StatusCheck:"",
      chksignOffStatus:false,
      hideAlert:true,
      isChecked:false,
      isreadOnly:true,
      isSignOffsummary:"",
      createdBy:"" ,
      CheckedSignOffArray:[] ,
      AllStatusSummary:"",
      signOffCount:"",
      ApproversList:"",
      SignOffList:"",
      Requester:"",
      RequestMail:"",
      RequesterJobTitle:"",
      RequestDate:"",
      RequesterDept:""
    }
    var currentURL = window.location.search.substring(1);
    var sURLVariables = currentURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) {
      var sParameterName = sURLVariables[i].split('=');
      if (sParameterName[0] == "SID") {
        this.state = { queryStringId: Number(sParameterName[1]) }
        this.queryStringId = Number(sParameterName[1]);
      }
      else if(sParameterName[0] == "CMode")
      {
        this.queryMode=sParameterName[1].toLowerCase();      
        this.state = { viewMode: true };
       
      }

    }
    if(this.queryMode=="editadmin")
    {
      this.state={headerContent:"Admin Sign-Offs"};
    }
    else{
      this.state={headerContent:"Self Sign-Offs"};
    }
    
    this.getSingOffDetails();
  }
  private submitSelfApprove = (event): void => {
  
  var UserName = this.props.spcontext.pageContext.user.displayName;
  var comments = "";
  var setSummary="";
  var date = new Date().toLocaleString();
  var lastSignOffSummary=this.state.isSignOffsummary;

  for(let i=0;i<signOffCheck.length;i++)
  {
    setSummary =setSummary+ signOffCheck[i] + "~SignOff Completed~SignOff Done by "+UserName+"~" + date + "|";

    if(lastSignOffSummary)
    var newSignOffSummary=lastSignOffSummary+"SignOffUser~"+signOffCheck[i]+"~"+"SignOff Completed"+"~"+date+"|"
    else
    var newSignOffSummary="SignOffUser~"+signOffCheck[i]+"~"+"SignOff Completed"+"~"+date+"|"
  }

  var finalSummary = this.state.AllStatusSummary + setSummary;
  var count = finalSummary.match(/SignOff Completed/g);
  var clean_count = !count ? false : count.length;
  var calcIdx = count.length;
  var SignOffUsersLength = this.state.getSignOffUser.length;

  var signoffAssignDate="";
  if (calcIdx == SignOffUsersLength) {
    var signOffUpdate = "SignOff Completed";
    var pendingAt="Sign-Off Stage";
    signoffAssignDate=new Date().toLocaleDateString();
  }
  else {
    var signOffUpdate = "SignOff Pending";
    var pendingAt="Approval Stage";
    signoffAssignDate=this.state.AssignedDate
  }
  var newsignOffCount=(parseInt(this.state.signOffCount)-1).toString()
    sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
      SignOffStatus:signOffUpdate,
      SignoffsSummary:newSignOffSummary,
      ApprovalSummary:finalSummary,
      SignOffUsersCount:newsignOffCount,
      PendingAt:pendingAt,
      AssignedDate:signoffAssignDate
    }).then((_sucess:any)=>{
      swal({
        title: 'Success',
        text: "Self Sign-Off Completed Successfully.!",
        icon: 'success',              
      }).then(()=>{
        window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
      });
      

  //  this.dialogContentProps = {
  //       title: "Self Sign-Off Completed Successfully.!",
  //     };
  //     this.setState({hideAlert:false});   
      console.log("Self approved done by requester")
    });
    
  }
  public cancelForm = (event): void => {
    window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
  }
  public toggleHideDialog = (event): void => {
    this.setState({hideAlert:true});
   // window.location.reload();
    window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
  }

  private _onChange(ev: React.FormEvent<HTMLInputElement>, isChecked: boolean) {
    console.log(ev.currentTarget.title, isChecked);
   if(isChecked )
    {
      signOffCheck.push(ev.currentTarget.title);
    }
    else
    {
      var index = signOffCheck.indexOf(ev.currentTarget.title);
      if(index!=-1){

        signOffCheck.splice(index, 1);
      }
    }
  }
  public async getSingOffDetails() {

    this.currentUserEmail = this.props.spcontext.pageContext.user.email;
    //this.setState({headerContent:"Tailgate JBC - Draft"});
      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).select("*,TaskIdentifier,Description,Author/EMail,Author/Title,Author/Id,Approvers/Title,Approvers/EMail,Approvers/Id,Signoffs/EMail,Signoffs/Id,Signoffs/Title").expand("Author,Approvers,Signoffs").get().then((singleItem: any) => {
        if (singleItem.Approvers && singleItem.Approvers.length >= 0) {
          var approversArray = [];
          var approversIDArray = [];
          for (let i = 0; i < singleItem.Approvers.length; i++) {
            approversArray.push(singleItem.Approvers[i].EMail.split('@')[0]);
            approversIDArray.push(singleItem.Approvers[i].Id);
            this.approverstextArray.push(singleItem.Approvers[i].Title);
            // if(i+1<=singleItem.ApproversId.length)
            // this.setState({getApprovepeoplePicker_User:approversArray})  
          }
       

        }
        if (singleItem.Signoffs.length > 0) {
          var signOffarray = [];
          var signOffIDArray = [];
  
  
          singleItem.Signoffs.map((result, idx) => {
            signOffarray.push(result.Title);
            signOffIDArray.push(result.Id);
            this.signOfftextArray.push(result.Title)
          });
        }
  
  
        var folderPath = "TaskDocuments/" + this.queryStringId
        sp.web.getFolderByServerRelativeUrl(folderPath).files.select('*,ID').get().then((allFiles) => {
          console.log(allFiles);
          var fetchFiles = [];
          allFiles.map((singleFile) => {
            fetchFiles.push({ filename: singleFile.Name, files: singleFile.ServerRelativeUrl })
          });
          const loginName = "i:0#.f|membership|"+singleItem.Author.EMail;
          sp.profiles.getPropertiesFor(loginName).then((pro)=>{
            var userProperties = pro.UserProfileProperties;
            userProperties.map((data)=>{
              if(data.Key=="Department")
              this.setState({RequesterDept:data.Value});
              else if(data.Key=="SPS-JobTitle")
              this.setState({RequesterJobTitle:data.Value})

            })  
            console.log(pro);
          });
          this.setState({
            topicValue: singleItem.TaskIdentifier,
            descriptionValue: singleItem.Description,
            getApprovepeoplePicker_User: approversArray,
            getSignOffUser: signOffarray,
            AllStatusSummary:singleItem.ApprovalSummary,
            StatusSummary:singleItem.SignoffsSummary,
            allpeoplePicker_User: approversIDArray,
            allpeoplePicker2_User: signOffIDArray,
            filePickerResult: fetchFiles,
            fetchApprovers:singleItem.ApproversId,
            StatusCheck:singleItem.SignOffStatus,
            fetchSignOffUsers:singleItem.SignoffsId,
            addformHide:singleItem.SignOffStatus=="SignOff Completed"?true:false,
            isSignOffsummary:singleItem.SignoffsSummary,
            createdBy:singleItem.Author.EMail,
            signOffCount:singleItem.SignOffUsersCount,
            ApproversList:this.approverstextArray.toString(),
            SignOffList:this.signOfftextArray.toString(),
            Requester:singleItem.Author.Title,
            RequestMail:singleItem.Author.EMail,
            RequestDate:new Date(singleItem.Created).toLocaleDateString()
          })
        });
      });
    }
    
  public render(): React.ReactElement<IReportSignOffProps> {
    var StatusSummary=[];
    var signOffUserArray=[];
    if(this.state.AllStatusSummary)
     StatusSummary=this.state.AllStatusSummary.split('|');
     if( this.state.getSignOffUser)
     signOffUserArray=this.state.getSignOffUser
    return (
      <div className={ styles.reportSignOff }>
          <h2>{this.state.headerContent}</h2>
          {/* <h2>Self Sign-Offs</h2> */}
        <div className={ styles.container }>    
        <div className="item-contents">            
        <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Topic </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.topicValue}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Description </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.descriptionValue}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Requester </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.Requester}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Request Date </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.RequestDate}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Requester Email </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.RequestMail}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Job Title </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.RequesterJobTitle}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Department </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.RequesterDept}</label>
                </div>
              </div>
              {
                this.state.ApproversList?<div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Approvers </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.ApproversList}</label>
                </div>
              </div>:""
              }
              

              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Sign Off's </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  <label>{this.state.SignOffList}</label>
                </div>
              </div>
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Attachments  </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
                <div className={styles.col_6}>
                  {this.state.filePickerResult && this.state.filePickerResult.map((filedet) => {
                    return (
                      <div>
                        <Link href={filedet.files}>{filedet.filename}</Link>

                      </div>
                    );
                  })}
                </div>
              </div>
              </div>
              <hr />

              <div className="history-contents">
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_6}>
                  <label className={styles.divalign}>Action History</label>
                </div>
                <div className={classnames(styles.col_12, "tbl-margin")}> 
                  <table className={styles.table}><thead><tr><th>Name</th><th>Action</th><th>Comments</th><th>Date</th></tr></thead><tbody>
                    {StatusSummary.length > 0 && StatusSummary.map((rowDet) => {
                      if (rowDet) {
                        rowDet = rowDet.split('~');
                        return (
                          <tr><td>{rowDet[0]}</td><td>{rowDet[1]}</td><td>{!rowDet[2] ? "-" : rowDet[2]}</td><td>{rowDet[3]}</td></tr>
                        );
                      }

                    })}
                  </tbody></table>
                </div>
              </div>
            </div>
<hr/>
          {/* <div className={styles.row}>
            <div className={styles.col_3}>
              <label>Attachments  </label>
            </div>
            <div className={styles.col_1}>
              :
            </div>
            <div className={styles.col_6}>
              {this.state.filePickerResult&&this.state.filePickerResult.map((filedet) => {
                return (
                  <div>
                    <Link href={filedet.files}>{filedet.filename}</Link>
                  </div>
                );  
              })}
            </div>
          </div> */}
          <div className="history-contents">
          <div className={styles.row}>
          <table className={classnames(styles.table, 'reportTable')}><thead><tr><th>Name</th><th>Date</th><th>Action</th></tr></thead><tbody>
                {
                   signOffUserArray.length>0&&signOffUserArray.map((signOffItem) => {
                    var ItemDet=signOffItem;
                    //  for(let i=0;i< StatusSummary.length;i++)
                    //  {
                       var rowDet=this.state.isSignOffsummary;
                       if(rowDet)
                       {
                         var existIdx=rowDet.indexOf(ItemDet)
                         if(existIdx>=0)
                         {
                          var getdatearr:any=[]
                          //  rowDet = rowDet.split('~');
                          var index= existIdx+ItemDet.length+18;
                          var getdate=rowDet.slice(index);
                           getdatearr=getdate["split"]('|')
                           return (
                             <tr><td>{ItemDet}</td><td>{getdatearr[0]}</td><td>{"Sign Off Completed"}</td></tr>
                           );
                         }
                         else{
                           return (
                             <tr><td>{ItemDet}</td><td>{"-"}</td><td>{<Checkbox label="Sign Off Pending"  title={ItemDet} onChange={this._onChange} />}</td></tr>
                           );
                         }
                       }
                       else{
                        return (
                          <tr><td>{ItemDet}</td><td>{"-"}</td><td>{<Checkbox label="Sign Off Pending"  title={ItemDet} onChange={this._onChange} />}</td></tr>
                        );
                      }
                    //  }
                   })
              }
              </tbody></table>
          </div>
          </div>
           <div className={classnames(styles.row, "reportBtn")}>
          <div className={classnames(styles.col_3, styles.btnCancel)}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />

            </div>
            <div className={styles.col_3}>         
            {/* hidden={this.state.addformHide}    */}
              <PrimaryButton className={styles.btnSubmit} text={this.state.headerContent} onClick={this.submitSelfApprove} />
            </div>
            </div>
        </div>
        {/* <div>   <Dialog
        hidden={this.state.hideAlert}  
        onDismiss={this.toggleHideDialog}    
        dialogContentProps={this.dialogContentProps}      
      >
        <DialogFooter>
          <PrimaryButton onClick={this.toggleHideDialog} text="Ok" />      
        </DialogFooter></Dialog>
        </div> */}
      </div>
    );
  }
  
}
