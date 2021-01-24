import * as React from 'react';
import styles from './AddTailGateRequest.module.scss';
import { IAddTailGateRequestProps } from './IAddTailGateRequestProps';
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
import { IconButton, IIconProps, IContextualMenuProps, Stack, Link,Modal } from 'office-ui-fabric-react';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import {
  Dialog,
  DialogFooter, 
  DialogType,
  IDialogStyles,
} from "office-ui-fabric-react/lib/Dialog";     
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
// import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { IAddTailGateRequestState } from './IAddTailGateRequestState';
import classnames from 'classnames';
import swal from 'sweetalert';
import '../../../ExternalRef/style.css';
export default class AddTailGateRequest extends React.Component<IAddTailGateRequestProps, IAddTailGateRequestState> {
  public state;
  public queryStringId: any;
  public queryMode: any;
  public summaryStatus: any;
  options: IChoiceGroupOption[];
  public approverstextArray=[];
  public signOfftextArray=[];

  public dialogContentProps = {
    //subText: '',
    title: '',
  };
  
  constructor(props: IAddTailGateRequestProps) {
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
      errorApproverTxt:"",
      errorSignOffTxt:"",
      errortopicValue: "",
      errordescriptionValue: "",
      errorfileAttach: "",
      errorapproverUsers: false,
      errorSignoffUsers: false,
      addformHide: false,
      getApprovepeoplePicker_User: [],
      getSignOffUser: [],
      queryStringId: 0,
      headerContent: "Tailgate JBC - New",
      viewMode: false,
      StatusSummary: "",
      fetchSignOffUsers: [],
      removedFiles: [],
      newFiles: [],
      fileDetails: [],
      approveStatus: "Approve",
      comments: "",
      errorcomments: "",
      fetchApprovers: [],
      StatusCheck: "",
      chksignOffStatus: false,
      hideAlert: true,
      singOffUserText: [],
      approversUserText: [],
      signoffSummary: "",
      signOffCount:"",
      AssignedDate:"",
      ApproversList:"",
      SignOffList:"",
      pendingAtStage:"",
      adminSelct:true,
      adminSelctApprover:"",
      errorDropDown:"",
      isGroupMember:false,
      adminApproverid:"",
      Requester:"",
      RequestMail:"",
      RequesterJobTitle:"",
      RequestDate:"",
      RequesterDept:"",
      modalApprovalContent:"",
      signOffModalContent:"",
      isApproverModalOpen:false,
      hideApproverModal:false
      // subText:""

    }
    this.options = [
      { key: 'A', text: 'Approve' },
      { key: 'B', text: 'Return' }
    ];
    var currentURL = window.location.search.substring(1);
    var sURLVariables = currentURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) {
      var sParameterName = sURLVariables[i].split('=');
      if (sParameterName[0] == "RID") {
        this.state = { queryStringId: Number(sParameterName[1]), headerContent: "Tailgate JBC - Draft" }
        this.queryStringId = Number(sParameterName[1]);
      }
      else if (sParameterName[0] == "CMode") {
        this.queryMode = sParameterName[1].toLowerCase();
        this.state = { viewMode: true, headerContent: "Tailgate JBC - View" }
      }

    }
    if(this.queryStringId)
    {
      this.init();

      this.getGroupDetails();
    }

  }

  DeleteIcon: IIconProps = { iconName: 'Delete' };
  UploadIcon:IIconProps = { iconName: 'BulkUpload' };
   cancelIcon: IIconProps = { iconName: 'Cancel' };

public init=()=>{
 
  sp.web.currentUser.get().then((UserId) => {
    this.getDraftDetails(UserId);
  })
}
public getGroupDetails=()=>{
  let groups = sp.web.currentUser.groups().then((grpDetails)=>{
    grpDetails.map((eachGroup,idx)=>{
      if(eachGroup.Title=="AdminTeam")
      {
        this.setState({isGroupMember:true});
        this.state.isGroupMember=true;
      }
    });
  })
}
  public Approverpeoplechange = (event) => {

    if (event["length"] > 0) {
      var resultarray = event.map((user) => user.id)

      this.setState({ allpeoplePicker_User: resultarray, errorapproverUsers: false });
      var resultUserarray = event.map((userText) => userText.text)
      this.setState({ approversUserText: resultUserarray });

    }
    else {
      this.setState({ allpeoplePicker_User: [], errorapproverUsers: true })
    }

  }

  public SignOffpeoplechange = (event) => {

    if (event["length"] > 0) {
      var resultarray = event.map((user) => user.id)
      this.setState({ allpeoplePicker2_User: resultarray, errorSignoffUsers: false });
      var resultUserarray = event.map((userText) => userText.text)
      this.setState({ singOffUserText: resultUserarray });
    }
    else {
      this.setState({ allpeoplePicker2_User: []})

    }

  }

  public cancelForm = (event): void => {
    window.location.href = this.props.siteURL + "/SitePages/TailGateRequestDashBoard.aspx";

  }
  private submitForm = (event): void => {
    var submitType = event.target.textContent == "Submit" ? "Submit" : "Draft";

    let approveinputValue = (document.querySelectorAll(".approve-Input, .ms-BasePicker-input")[1] as HTMLInputElement).value;
    let signoffinputValue = (document.querySelectorAll(".SignOff-Input, .ms-BasePicker-input")[2] as HTMLInputElement).value;
    // errorApproverTxt:"",
    // errorSignOffTxt:"",
    (approveinputValue)?this.setState({errorApproverTxt:"User doesn't existed",errorapproverUsers:true}):"";
    (signoffinputValue)?this.setState({errorSignOffTxt:"User doesn't existed",errorSignoffUsers: true}):"";



    // !this.state.isGroupMember&&this.state.allpeoplePicker_User.length==0&&!approveinputValue?this.setState({errorApproverTxt:"Approver is required",errorapproverUsers:true}):this.setState({errorApproverTxt:"User doesn't existed",errorapproverUsers:true}):this.setState({errorApproverTxt:"Approver is required",errorapproverUsers:true});

    // this.state.allpeoplePicker2_User.length==0 && !signoffinputValue?  this.setState({ errorSignOffTxt:"SignOff User is required",errorSignoffUsers: true }):this.setState({errorSignOffTxt:"User doesn't existed",errorSignoffUsers: true});

    if(submitType=="Submit")
    {
      this.state.topicValue ? "" : this.setState({ errortopicValue: "Topic is required" });
      this.state.descriptionValue ? "" : this.setState({ errordescriptionValue: "Description is required" });
      this.state.filePickerResult.length == 0 ? this.setState({ errorfileAttach: "Attachments are required" }) : "";

      if(!this.state.isGroupMember&&this.state.allpeoplePicker_User.length==0&&!approveinputValue)
      this.setState({errorApproverTxt:"Approver is required",errorapproverUsers:true})
      else if(approveinputValue)
      this.setState({errorApproverTxt:"User doesn't existed",errorapproverUsers:true})
      else
      this.setState({errorApproverTxt:"",errorapproverUsers:false})
  
      if(this.state.allpeoplePicker2_User.length==0 && !signoffinputValue)
      this.setState({ errorSignOffTxt:"SignOff User is required",errorSignoffUsers: true })
      else if(signoffinputValue)
      this.setState({errorSignOffTxt:"User doesn't existed",errorSignoffUsers: true})
      else
      this.setState({errorSignOffTxt:"",errorSignoffUsers: false})

    }
    else
    {
      this.state.topicValue.trim() ? "" : this.setState({ errortopicValue: "Topic is required" });

      if(approveinputValue)
      this.setState({errorApproverTxt:"User doesn't existed",errorapproverUsers:true})
      else
      this.setState({errorApproverTxt:"",errorapproverUsers:false})

      if(signoffinputValue)
      this.setState({errorSignOffTxt:"User doesn't existed",errorSignoffUsers: true})
      else
      this.setState({errorSignOffTxt:"",errorSignoffUsers: false})

    }
    if (!this.queryStringId) {
      if((submitType=="Draft"&&this.state.topicValue&&!this.state.errorSignoffUsers&&!this.state.errorapproverUsers)||(submitType=="Submit"&&this.state.isGroupMember&& (this.state.topicValue && this.state.descriptionValue && this.state.filePickerResult.length > 0 && this.state.allpeoplePicker2_User.length > 0))||(submitType=="Submit"&&!this.state.isGroupMember&& (this.state.topicValue && this.state.descriptionValue && this.state.filePickerResult.length > 0 && this.state.allpeoplePicker2_User.length > 0&&this.state.allpeoplePicker_User.length>0)))
      {
        let today = new Date().toISOString().slice(0, 10);
        var UserName = this.props.spcontext.pageContext.user.displayName;
        var comments = this.state.descriptionValue;
        var date = new Date().toLocaleString();
        var setSummary = UserName + "~New Request~-~" + date + "|";
        var usersSignoff = '';
        var approverFlow = this.state.allpeoplePicker_User? this.state.allpeoplePicker_User.toString():"0"
        var signOffFlow =this.state.allpeoplePicker2_User? this.state.allpeoplePicker2_User.toString():"0"
         if (this.state.singOffUserText && this.state.singOffUserText.length > 0 ) {
          for (var i = 0; i < this.state.singOffUserText.length; i++) {
            usersSignoff = usersSignoff +  "~SignOffUser~" + this.state.singOffUserText[i] + "~" + "Pending" + "~" + date + "|";
          }
        }
       document.getElementById("loader-container").style.display= 'flex'; 
        sp.web.lists.getByTitle("TailgateTasksActivity").items.add({
          Title: "Tailgate",
          Description: this.state.descriptionValue,
          Status: event.target.textContent == "Submit" ? "Submit" : "Draft",
          ApprovalStatus: event.target.textContent == "Submit" && !approverFlow ? "Approved" : "",
          SignOffStatus:event.target.textContent == "Submit" && !approverFlow ? "SignOff Pending" : "",
          ProcessType: "Tailgate",
          TaskIdentifier: this.state.topicValue,
          ApproversId: {
            results: this.state.allpeoplePicker_User && this.state.allpeoplePicker_User.length > 0 ? this.state.allpeoplePicker_User : [] 
          },
          SignoffsId: {
            results: this.state.allpeoplePicker2_User && this.state.allpeoplePicker2_User.length > 0 ? this.state.allpeoplePicker2_User : []  
          },
          ApprovalSummary: setSummary,
        ApproversCount:this.state.allpeoplePicker_User ?this.state.allpeoplePicker_User.length.toString():"0",
          SignOffUsersCount:this.state.allpeoplePicker2_User ?this.state.allpeoplePicker2_User.length.toString():"0",
          ApproversText:approverFlow,
          SignOffUsersText:signOffFlow,
          IsSubmitted:event.target.textContent == "Submit" ? "Yes" : "No"  ,
          ApprovalProcessStart:"No",
          PendingAt:event.target.textContent == "Submit"&& approverFlow? "Approval Stage" : event.target.textContent == "Submit"&& !approverFlow? "Sign-Off Stage":"",
          AssignedDate:event.target.textContent == "Submit" ? new Date().toLocaleDateString()
          : "",
          AdminApprovalSummary:""
        })
          .then((disID) => {
            sp.web.getFolderByServerRelativeUrl("TaskDocuments").folders.add("TaskDocuments" + '/' + disID.data.Id).then(result => {
              var allUploadFiles = this.state.filePickerResult;
              var tobeRemove = this.state.removedFiles;
              if (tobeRemove) {
                tobeRemove.map((re) => {
                  sp.web.getFileByServerRelativeUrl(re.files).recycle().then(() => {
                    console.log("deleted")
                  });
                });
              }
              if (allUploadFiles.length > 0) {
                this.EachfileUpload(allUploadFiles, result, disID.data.Id, submitType);
              }
              else {
                document.getElementById("loader-container").style.display= 'none';
                this.setState({
                  topicValue: "",
                  descriptionValue: "",
                  filePickerResult: [],
                  allpeoplePicker_User: [],
                  allpeoplePicker2_User: []
                });
                console.log("File upload successfully...!");
                document.getElementById("loader-container").style.display= 'none';
                swal({
                  title: 'Success',
                  text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
                  icon: 'success',              
                }).then(()=>{
                  window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
                });
              }
            });
          });
      }

        
    }
    else {
      if((submitType=="Draft"&&this.state.topicValue&&!this.state.errorSignoffUsers&&!this.state.errorapproverUsers)||(submitType=="Submit"&&this.state.isGroupMember&& (this.state.topicValue && this.state.descriptionValue && this.state.filePickerResult.length > 0 && this.state.allpeoplePicker2_User.length > 0))||(submitType=="Submit"&&!this.state.isGroupMember&& (this.state.topicValue && this.state.descriptionValue && this.state.filePickerResult.length > 0 && this.state.allpeoplePicker2_User.length > 0&&this.state.allpeoplePicker_User.length>0)))
      {
        let today = new Date().toISOString().slice(0, 10);
        var UserName = this.props.spcontext.pageContext.user.displayName;
        var comments = this.state.descriptionValue;
        var date = new Date().toLocaleString();
        var setSummary = UserName + "~New Request~-~" + date + "|";
        var approverFlow = this.state.allpeoplePicker_User? this.state.allpeoplePicker_User.toString():"0"
        var signOffFlow =this.state.allpeoplePicker2_User? this.state.allpeoplePicker2_User.toString():"0"
        var usersSignoff = '';
        var usersApprovers = '';
        if (this.state.singOffUserText && this.state.singOffUserText.length > 0 ) {
          for (var i = 0; i < this.state.singOffUserText.length; i++) {
            usersSignoff = usersSignoff +  "~SignOffUser~" + this.state.singOffUserText[i] + "~" + "Pending" + "~" + date + "|";
          }
        }
        document.getElementById("loader-container").style.display= 'flex'; 
        sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
          Title: "Tailgate",
          Description: this.state.descriptionValue,
          Status: event.target.textContent == "Submit" ? "Submit" : "Draft",
          ApprovalStatus: event.target.textContent == "Submit" && !approverFlow ? "Approved" : "",
          SignOffStatus:event.target.textContent == "Submit" && !approverFlow ? "SignOff Pending" : "",
          ProcessType: "Tailgate",
          TaskIdentifier: this.state.topicValue,
          ApproversId: {
            results: this.state.allpeoplePicker_User && this.state.allpeoplePicker_User.length > 0 ? this.state.allpeoplePicker_User : []  
          },
          SignoffsId: {
            results: this.state.allpeoplePicker2_User.length > 0 ? this.state.allpeoplePicker2_User : []  // User/ 
          },
          ApprovalSummary: setSummary,
          ApproversCount:this.state.allpeoplePicker_User ?this.state.allpeoplePicker_User.length.toString():"0",
          SignOffUsersCount:this.state.allpeoplePicker2_User ?this.state.allpeoplePicker2_User.length.toString():"0",
          ApproversText:approverFlow,
          SignOffUsersText:signOffFlow,
          IsSubmitted:event.target.textContent == "Submit" ? "Yes" : "No",
          ApprovalProcessStart:"No",
          PendingAt:event.target.textContent == "Submit"&& approverFlow? "Approval Stage" : event.target.textContent == "Submit"&& !approverFlow? "Sign-Off Stage":"",
          AssignedDate:event.target.textContent == "Submit" ? new Date().toLocaleDateString()
          : "",
          AdminApprovalSummary:""
         
        })
          .then((draftId) => {
            sp.web.getFolderByServerRelativeUrl("TaskDocuments").folders.add("TaskDocuments" + '/' + this.queryStringId).then(result => {
              var allUploadFiles = this.state.filePickerResult;
              var tobeRemove = this.state.removedFiles;
              if (tobeRemove) {
                tobeRemove.map((re) => {
                  sp.web.getFileByServerRelativeUrl(re.files).recycle().then(() => {
                    console.log("deleted")
                  });
                });
              }
              if (allUploadFiles.length > 0) {
                this.EachfileUpload(allUploadFiles, result, this.queryStringId, submitType);
              }
              else {
                console.log("File upload successfully...!");
  
                this.setState({
                  topicValue: "",
                  descriptionValue: "",
                  filePickerResult: [],
                  allpeoplePicker_User: [],
                  allpeoplePicker2_User: []
                });
                document.getElementById("loader-container").style.display= 'none';
                swal({
                  title: 'Success',
                  text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
                  icon: 'success',              
                }).then(()=>{
                  window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
                });         
              }
            });
          });
      }


    }

  }
  async EachfileUpload(allUploadFiles, result, newId, submitType) {
    // await allUploadFiles.map((eachfileDetails, index) => {
    for (let i = 0; i < allUploadFiles.length; i++) {
      if (allUploadFiles[i].files["name"]) {
        await result.folder.files.add(allUploadFiles[i].filename, allUploadFiles[i].files, true).then((fresult) => {
          if (allUploadFiles.length <= i + 1) {
            this.setState({
              topicValue: "",
              descriptionValue: "",
              filePickerResult: [],
              allpeoplePicker_User: [],
              allpeoplePicker2_User: []
            });
            console.log("File upload successfully...!");
            document.getElementById("loader-container").style.display= 'none'; 
            swal({
              title: 'Success',
              text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
              icon: 'success',              
            }).then(()=>{
              window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
            });

          }
        });
      }
      else {
        if (allUploadFiles.length <= i + 1) {
          this.setState({
            topicValue: "",
            descriptionValue: "",
            filePickerResult: [],
            allpeoplePicker_User: [],
            allpeoplePicker2_User: []
          })

          console.log("File upload successfully...!");
          swal({
            title: 'Success',
            text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
            icon: 'success',              
          }).then(()=>{
            window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
          });

        }
      }
    }
  }

  public fileUploadCallback = (e) => {
    if (!this.queryStringId) {
      var files = e.target.files;
      var isExist=true;
    if(files[0])
    {
      var fname=files[0].name.toLowerCase();
      const ext = ['.jpg', '.jpeg', '.png', '.xls', '.xlsx', '.ppt','.pptx','.doc','.docx','pdf'];
      var isExist= ext.some(el => fname.endsWith(el));
  }
      
      if (files && files.length > 0 && files[0].size <= 5000000&&isExist) {
        


        var allfiles = [];
        for (let i = 0; i < files.length; i++) {
          if (this.state.filePickerResult)
            allfiles = this.state.filePickerResult;
          var sepArray = allfiles.filter((eleFile) => { return eleFile.filename == files[i].name })
          if (sepArray.length <= 0) {
            allfiles.push({ filename: files[i].name, files: files[i] });
          }
          if (files.length <= i + 1)
            this.setState({ filePickerResult: allfiles, errorfileAttach: "" });
        }

        e.target.value = null;

      }
      else if(!isExist)
      {
        this.setState({errorfileAttach: "The system does not support this file type" });
        e.target.value = null;
      }
      else if (files && files.length > 0 && files[0].size >= 5000000) {
        this.setState({  errorfileAttach: "Maximum attachment size 5 MB" });
        e.target.value = null;
      }
      else {
        this.setState({ filePickerResult: []});

        e.target.value = null;
      }
    }
    else {
      var files = e.target.files;
      var isExist=true;
      if(files[0])
      {
        var fname=files[0].name.toLowerCase();
        const ext = ['.jpg', '.jpeg', '.png', '.xls', '.xlsx', '.ppt','.pptx','.doc','.docx','pdf'];
        var isExist= ext.some(el => fname.endsWith(el));
     }

      if (files && files.length > 0&& files[0].size <= 5000000&&isExist) {
        var allfiles = [];
        var oldallfiles = [];
        for (let i = 0; i < files.length; i++) {
          if (this.state.newFiles)
            allfiles = this.state.newFiles;
          if (this.state.filePickerResult)
            oldallfiles = this.state.filePickerResult;
          var sepArray = allfiles.filter((eleFile) => { return eleFile.filename == files[i].name })
          if (sepArray.length <= 0) {
            allfiles.push({ filename: files[i].name, files: files[i] });
            oldallfiles.push({ filename: files[i].name, files: files[i] })
          }

          if (files.length <= i + 1)
            this.setState({ newFiles: allfiles, filePickerResult: oldallfiles });
        }
        e.target.value = null;
      }
      else if(!isExist)
      {
        this.setState({errorfileAttach: "The system does not support this file type" });
        e.target.value = null;
      }
      else if (files && files.length > 0 && files[0].size >= 5000000) {
        this.setState({  errorfileAttach: "Maximum attachment size 5 MB" });
        e.target.value = null;
      } else {
        this.setState({ newFiles: [] });

        e.target.value = null;
      }
    }
  }

  public async getDraftDetails(UserId) {
    this.setState({approveStatus:"Approve",adminSelct:true});
    sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).select("*,TaskIdentifier,Description,Approvers/EMail,Approvers/Id,Approvers/Title,Signoffs/EMail,Signoffs/Id,Signoffs/Title,Author/Title,Author/EMail").expand("Approvers,Signoffs,Author").get().then((singleItem: any) => {
      var approversArray = [];
      var approversIDArray = [];
      var showApprovers=[];
      var signOffarray = [];
      var signOffIDArray = [];
      var modalApproverContent="";
      var signOffModalContent="";
      if (singleItem.Approvers && singleItem.Approvers.length >= 0) {

        for (let i = 0; i < singleItem.Approvers.length; i++) {
          approversArray.push(singleItem.Approvers[i].EMail);
          approversIDArray.push(singleItem.Approvers[i].Id);
          showApprovers.push(singleItem.Approvers[i].Title)
          this.approverstextArray.push({Title:singleItem.Approvers[i].Title,ID:singleItem.Approvers[i].Id});
          if(singleItem.Approvers.length>5)
          {
            modalApproverContent+="<p>"+singleItem.Approvers[i].Title+"</p>"
          }
        }


      }
      if (singleItem.Signoffs&&singleItem.Signoffs.length > 0) {
 


        singleItem.Signoffs.map((result, idx) => {
          signOffarray.push(result.EMail);
          signOffIDArray.push(result.Id);
          this.signOfftextArray.push(result.Title)
          if(singleItem.Signoffs.length>5)
          {
            signOffModalContent+="<p>"+result.Title+"</p>"
          }
        });
      }
      var responseIdx=0;
      var signresponseIdx=0
      if(singleItem.ApproversId)
      {
        var approveIdx=singleItem.ApproversId.includes(UserId['Id'])
      }
      if(singleItem.SignoffsId)
      {
        var signIdx=singleItem.SignoffsId.includes(UserId['Id'])
      }
      var userDisplayName=this.props.spcontext.pageContext.user.displayName;
      var checkCondition=userDisplayName+"~Approved";
      responseIdx=singleItem.ApprovalSummary.indexOf(checkCondition);
      var checkSignoff=userDisplayName+"~SignOff Completed";
      signresponseIdx=singleItem.ApprovalSummary.indexOf(checkSignoff);

      if((responseIdx>=0||signresponseIdx>=0)&&this.queryMode=="edit")
      {
        swal({
          title: 'Information',
          text: "You Have already updated this request",
          icon: 'error',              
        }).then(()=>{
          window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
        });
      }
      else
      {
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
            StatusSummary: singleItem.ApprovalSummary,
            signoffSummary:singleItem.SignoffsSummary,
            allpeoplePicker_User: approversIDArray,
            allpeoplePicker2_User: signOffIDArray,
            filePickerResult: fetchFiles,
            fetchApprovers: singleItem.ApproversId,
            StatusCheck: singleItem.SignOffStatus,
            fetchSignOffUsers: singleItem.SignoffsId,
            signOffCount:singleItem.SignOffUsersCount,
            AssignedDate:singleItem.AssignedDate,
            approveStatus:"Approve",
            ApproversList:showApprovers.toString(),
            SignOffList:this.signOfftextArray.toString(),
            pendingAtStage:singleItem.PendingAt,
            ApproverAdminSummary:singleItem.AdminApprovalSummary,
            Requester:singleItem.Author.Title,
            RequestMail:singleItem.Author.EMail,
            RequestDate:new Date(singleItem.Created).toLocaleDateString(),
            modalApprovalContent:modalApproverContent,
            signOffModalContent:signOffModalContent
          })
        });
      }


    });
  }

  public removeDoc = (e) => {

    var targetelement;
    if (!this.queryStringId) {
      targetelement = e.currentTarget.id;
      var filesArray = this.state.filePickerResult;
      filesArray = filesArray.filter((key, index) => { return index != targetelement; });

      this.setState({ filePickerResult: filesArray });
    }
    else {
      targetelement = parseInt(e.currentTarget.id);
      var filesArray = this.state.filePickerResult;
      var removedFiles = this.state.removedFiles;
      removedFiles = filesArray.filter((key, index) => {
        return index == targetelement
      });

      filesArray = filesArray.filter((key, index) => {
        return index != targetelement
      });
      this.setState({ filePickerResult: filesArray, removedFiles: removedFiles });
    }

  }
  public _alertClicked = (): void => {
    this.setState({
      topicValue: "",
      descriptionValue: "",
      fileDetails: []
    });
    window.location.href = this.props.siteURL + "/SitePages/TailGateRequestDashBoard.aspx";
  }

  public AdminApprovalForm =(e) =>{
    var approvingId=this.state.adminApproverid.toString();

    if(!this.state.adminSelctApprover)
    {
      this.setState({errorDropDown:"Please select approver"});
      return false
    }
    var adminName=this.props.spcontext.pageContext.user.displayName;
    if (this.state.approveStatus == "Approve") {
     
      var UserName = this.state.adminSelctApprover;
      var comments = this.state.comments;
      var date = new Date().toLocaleString();
      if (comments)
        var setSummary = UserName + "~Approved~Approved done by "+adminName+" " + comments + "~" + date + "|";
      else
        var setSummary = UserName + "~Approved~" + "Approved done by "+adminName + "~" + date + "|";

      var finalSummary = this.state.StatusSummary + setSummary;
      var finalIdx = finalSummary.split('|');
      var calcIdx = finalIdx.length - 2;
      var signoffAssignDate="";
      var approversLength = this.state.fetchApprovers.length;
      if (calcIdx == approversLength) {
        var statusUpdate = "Approved";
        var signOffUpdate = "SignOff Pending";
        var approverstatus = "Approved";
        var pendingAt="Sign-Off Stage";
         signoffAssignDate=new Date().toLocaleDateString();
      }
      else {
        var statusUpdate = "Submit";
        var approverstatus = "Approved";
        var pendingAt="Approval Stage";
         signoffAssignDate=this.state.AssignedDate
      }
      var lastAdminApprovalSummary=this.state.ApproverAdminSummary;
      if(lastAdminApprovalSummary)
      var newAdminSummary=lastAdminApprovalSummary+"ApproverUser~"+UserName+"~"+"Approved"+"~"+date+"|"
      else
      var newAdminSummary="ApproverUser~"+UserName+"~"+"Approved"+"~"+date+"|"
      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
        ApprovalSummary: this.state.StatusSummary + setSummary,
        Status: statusUpdate,
        SignOffStatus: signOffUpdate,
        ApprovalStatus: approverstatus,
        ApprovalProcessStart:"Yes",
        PendingAt:pendingAt,
        AssignedDate:signoffAssignDate,
        AdminApprovalSummary:newAdminSummary,
        AdminApprovingId:approvingId
      }).then(s => {
        swal({
          title: 'Success',
          text: "Request updated successfully..!!",
          icon: 'success',              
        }).then(()=>{
          window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
        });

      });
    } else if (this.state.approveStatus == "Return") {
      if (!this.state.comments) {
        this.setState({ errorcomments: "Comments is Required" });

      }
      else {
        var UserName = this.state.adminSelctApprover;
        var comments = this.state.comments;
        var date = new Date().toLocaleString();
        var setSummary = UserName + "~Returned~Returned Done by "+adminName + comments + "~" + date + "|";
        var statusUpdate = "Returned";
        var lastAdminApprovalSummary=this.state.ApproverAdminSummary;

      if(lastAdminApprovalSummary)
      var newAdminSummary=lastAdminApprovalSummary+"ApproverUser~"+UserName+"~"+" Returned"+"~"+date+"|"
      else
      var newAdminSummary="ApproverUser~"+UserName+"~"+"Returned"+"~"+date+"|"

        sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
          ApprovalSummary: this.state.StatusSummary + setSummary,
          Status: statusUpdate,
          ApprovalStatus: statusUpdate,
          ApprovalProcessStart:"Yes",
          AdminApprovalSummary:newAdminSummary,
          AdminApprovingId:approvingId
        }).then(s => {
          swal({
            title: 'Success',
            text: "Request updated successfully..!!",
            icon: 'success',              
          }).then(()=>{
            window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
          });

        });
      }

    }
  }
  public ApprovalSubmitForm = (): void => {
    if (this.state.approveStatus == "Approve") {
      var UserName = this.props.spcontext.pageContext.user.displayName;
      var comments = this.state.comments;
      var date = new Date().toLocaleString();
      if (comments)
        var setSummary = UserName + "~Approved~" + comments + "~" + date + "|";
      else
        var setSummary = UserName + "~Approved~" + "-" + "~" + date + "|";

      var finalSummary = this.state.StatusSummary + setSummary;
      var finalIdx = finalSummary.split('|');
      var calcIdx = finalIdx.length - 2;
      var signoffAssignDate="";
      var approversLength = this.state.fetchApprovers.length;
      if (calcIdx == approversLength) {
        var statusUpdate = "Approved";
        var signOffUpdate = "SignOff Pending";
        var approverstatus = "Approved";
        var pendingAt="Sign-Off Stage";
         signoffAssignDate=new Date().toLocaleDateString();
      }
      else {
        var statusUpdate = "Submit";
        var approverstatus = "Approved";
        var pendingAt="Approval Stage";
         signoffAssignDate=this.state.AssignedDate
      }
      var lastAdminApprovalSummary=this.state.ApproverAdminSummary;
      if(lastAdminApprovalSummary)
      var newAdminSummary=lastAdminApprovalSummary+"ApproverUser~"+UserName+"~"+"Approved"+"~"+date+"|"
      else
      var newAdminSummary="ApproverUser~"+UserName+"~"+"Approved"+"~"+date+"|"
      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
        ApprovalSummary: this.state.StatusSummary + setSummary,
        Status: statusUpdate,
        SignOffStatus: signOffUpdate,
        ApprovalStatus: approverstatus,
        ApprovalProcessStart:"Yes",
        PendingAt:pendingAt,
        AssignedDate:signoffAssignDate,
        AdminApprovalSummary:newAdminSummary
      }).then(s => {
        swal({
          title: 'Success',
          text: "Request updated successfully..!!",
          icon: 'success',              
        }).then(()=>{
          window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
        });
      });
    } else if (this.state.approveStatus == "Return") {
      if (!this.state.comments) {
        this.setState({ errorcomments: "Comments is Required" });

      }
      else {
        var UserName = this.props.spcontext.pageContext.user.displayName;
        var comments = this.state.comments;
        var date = new Date().toLocaleString();
        var setSummary = UserName + "~Returned~" + comments + "~" + date + "|";
        var statusUpdate = "Returned";
        var lastAdminApprovalSummary=this.state.ApproverAdminSummary;

      if(lastAdminApprovalSummary)
      var newAdminSummary=lastAdminApprovalSummary+"ApproverUser~"+UserName+"~"+"Returned"+"~"+date+"|"
      else
      var newAdminSummary="ApproverUser~"+UserName+"~"+"Returned"+"~"+date+"|"

        sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
          ApprovalSummary: this.state.StatusSummary + setSummary,
          Status: statusUpdate,
          ApprovalStatus: statusUpdate,
          ApprovalProcessStart:"Yes",
          AdminApprovalSummary:newAdminSummary
        }).then(s => {
          console.log("Items updated Successfully");
          swal({
            title: 'Success',
            text: "Request updated successfully..!!",
            icon: 'success',              
          }).then(()=>{
            window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
          });
        });
      }

    }

  }
  public toggleHideDialog = (event): void => {
    this.setState({ hideAlert: true });
  }
  public showApprover =()=>{
    const wrapper = document.createElement('div');
    wrapper.classList.add("allApproverList");
    wrapper.innerHTML = this.state.modalApprovalContent
swal({
  title: 'All Approvers',
  content: {
    element: wrapper,
  }

});
}

public showSignOffs =()=>{
  // this.setState({isApproverModalOpen:false});
  const wrapper = document.createElement('div');
  wrapper.classList.add("allApproverList");
  wrapper.innerHTML = this.state.signOffModalContent
//     var parent=document.getElementById('Approve-content');
// parent.insertAdjacentHTML('beforeend',this.state.modalApprovalContent)

swal({
title: 'All Sign-Offs',
content: {
  element: wrapper,
}

});
}
  

  public SignoffForm = () => {
    if (this.state.chksignOffStatus == true) {
      var UserName = this.props.spcontext.pageContext.user.displayName;
      var comments = "";
      var date = new Date().toLocaleString();
      if (comments)
        var setSummary = UserName + "~SignOff Completed~" + comments + "~" + date + "|";
      else
        var setSummary = UserName + "~SignOff Completed~" + "-" + "~" + date + "|";
      var finalSummary = this.state.StatusSummary + setSummary;
      var count = finalSummary.match(/SignOff Completed/g);
      var clean_count = !count ? false : count.length;
      var calcIdx = count.length;
      var SignOffUsersLength = this.state.fetchSignOffUsers.length;
      var lastSignOffSummary=this.state.signoffSummary;
      if(lastSignOffSummary)
      var newSignOffSummary=lastSignOffSummary+"SignOffUser~"+UserName+"~"+"SignOff Completed"+"~"+date+"|"
      else
      var newSignOffSummary="SignOffUser~"+UserName+"~"+"SignOff Completed"+"~"+date+"|"

      if (calcIdx == SignOffUsersLength) {
        var signOffUpdate = "SignOff Completed";
        var pendingAt="SignOff Completed";
      }
      else {
        var signOffUpdate = "SignOff Pending";
        var pendingAt="Sign-Off Stage";
      }
      var newsignOffCount=(parseInt(this.state.signOffCount)-1).toString()
      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.queryStringId).update({
        ApprovalSummary: this.state.StatusSummary + setSummary,
        SignOffStatus: signOffUpdate,
        SignoffsSummary: newSignOffSummary,
        SignOffUsersCount:newsignOffCount,
        PendingAt:pendingAt
      }).then(s => {
        swal({
          title: 'Success',
          text: "Request updated successfully..!!",
          icon: 'success',              
        }).then(()=>{
          window.location.href=this.props.siteURL+"/SitePages/TailGateRequestDashBoard.aspx";
        });
      });
    } else if (this.state.chksignOffStatus == "Return") {


    }
  }

  public render(): React.ReactElement<IAddTailGateRequestProps> {

    const options: IChoiceGroupOption[] = [
      { key: 'A', text: 'Approve' },
      { key: 'B', text: 'Return' }

    ];
    const Dropoptions: IDropdownOption[] = [

    ];
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 300 },
    };

    var StatusSummary = [];
    var appContent="";
    if (this.state.StatusSummary)
      StatusSummary = this.state.StatusSummary.split('|');
      if( this.state.modalApprovalContent)
      appContent=this.state.modalApprovalContent
    return (
 
      <><div>
        {
               
          !this.queryMode ? <><div id="loader-container" style={{display:"none"}}>
            <div className="loader">
              <span></span>
              <span></span>
              <span></span>
              <span></span>
              <span></span> 
            </div>
          </div><div className={styles.addTailGateRequest}>
              <h1>{this.state.headerContent}</h1>
              <div className={styles.container}>

                <div className={styles.row}>
                  <div className={classnames(styles.col_6, styles.fieldLbl)}>
                    <TextField label="Topic" required readOnly={this.state.viewMode}
                      value={this.state.topicValue}
                      onChanged={newVal => {
                        newVal && newVal.length > 100
                          ? this.setState({
                            topicValue: newVal,
                            errortopicValue: "Topic should not be more than 100 Characters"
                          })
                          : this.setState({
                            topicValue: newVal,
                             errortopicValue: ""
                          });
                      } }
                      errorMessage={this.state.errortopicValue} />
                  </div>
                  <div className={classnames(styles.col_6, styles.fieldLbl)}>
                    <Label required={this.state.isGroupMember?false:true}>Approvers</Label>

                    <PeoplePicker peoplePickerCntrlclassName="approve-Input"
                      context={this.props.spcontext}
                      titleText=""
                      personSelectionLimit={10}
                      groupName={""}
                      showtooltip={false}
                      // isRequired={true}
                      defaultSelectedUsers={this.state.getApprovepeoplePicker_User}
                      disabled={this.state.viewMode}
                      ensureUser={true}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      onChange={(e) => this.Approverpeoplechange.call(this, e)} />
  {this.state.errorapproverUsers ? <Label className={styles.pickerlabelErrormsg}>{this.state.errorApproverTxt}</Label> : ""}
                  </div>
                </div>
                <div className={styles.row}>
                  <div className={classnames(styles.col_6, styles.fieldLbl)}>
                    <TextField label="Description" required readOnly={this.state.viewMode}
                      value={this.state.descriptionValue}
                      onChanged={newDesVal => {
                        newDesVal && newDesVal.length > 3000
                          ? this.setState({
                            descriptionValue: newDesVal,
                            errordescriptionValue: "Description should not be more than 3000 Characters"
                          })
                          : this.setState({
                            descriptionValue: newDesVal,
                            errordescriptionValue: ""
                          });
                      } }
                      multiline rows={3} errorMessage={this.state.errordescriptionValue} />
                  </div>
                  <div className={classnames(styles.col_6, styles.fieldLbl)}>
                    <Label required>Sign offs</Label>
                    <PeoplePicker
                    peoplePickerCntrlclassName="SignOff-Input"
                      //  peoplePickerCntrlclassName={styles.pickerErrormsg}
                      context={this.props.spcontext}
                      titleText=""
                      personSelectionLimit={10}
                      groupName={""}
                      showtooltip={false}
                      //  isRequired={true}
                      defaultSelectedUsers={this.state.getSignOffUser}
                      disabled={this.state.viewMode}
                      ensureUser={true}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}  
                      resolveDelay={1000}
                      onChange={(e) => this.SignOffpeoplechange.call(this, e)} />
                    {this.state.errorSignoffUsers ? <Label className={styles.pickerlabelErrormsg}>{this.state.errorSignOffTxt}</Label> : ""}

                  </div>
                </div>    
                <div className={styles.row}> 
                  <div className={classnames(styles.col_6, styles.fieldLbl)}> 
                    <div>
                      <Label required>Attachments</Label>
                      {/* <div className="upload-group"> */}  
                        <div className="custom-upload">
                       
                          <input  id="fileCus" disabled={this.state.viewMode} type="file" multiple accept=".xlsx,.xls,.doc, image/*, .docx,.ppt, .pptx,.txt,.pdf,.png,.jpg" onChange={this.fileUploadCallback} />
                          <label htmlFor="fileCus" className="fileLbl"> <IconButton className={"file-upload-icon"} iconProps={this.UploadIcon}>  
                    </IconButton>Choose File</label>
                        </div>
                        {/* </div> */}
                      {this.state.filePickerResult ?
                        this.state.filePickerResult.map((filedet, index) => { 

                          return (
                            <div className={styles.attach}>
                              <label style={{ color: "#333" }}>{filedet.filename} </label>
                              <IconButton disabled={this.state.viewMode} className={styles.btntransparent} iconProps={this.DeleteIcon} onClick={this.removeDoc.bind(this)} id={index.toString()}>
                              </IconButton><br></br>
                            </div>
                          );

                        }) : ""}

                      {this.state.errorfileAttach ? <Label className={styles.pickerlabelErrormsg}>{this.state.errorfileAttach}</Label> : ""}
                    </div>

                  </div>
                </div>


                <div className={styles.row}>
                  <div className={classnames(styles.col_3, styles.btnAction)}>
                    <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />

                  </div>
                  <div className={classnames(styles.col_3, styles.btnAction)} hidden={this.state.viewMode}>
                    <PrimaryButton className={styles.btnDraft} text="Save as Draft" onClick={this.submitForm} />
                  </div>
                  <div className={classnames(styles.col_3, styles.btnAction)} hidden={this.state.viewMode}>
                    <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitForm} />
                  </div>
                </div>

              </div> 
            </div></> : this.queryMode == "edit" && this.queryStringId ? <div className={styles.addTailGateRequest}>
            <div className="container">
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
                this.state.fetchApprovers&&this.state.fetchApprovers.length>0?<div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Approvers </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
            {
              this.state.fetchApprovers.length<=5?<div className={styles.col_6}>
              <label>{this.state.ApproversList}</label>
            </div>:<Link onClick={this.showApprover}>All Approvers</Link>
            }

              </div>:""
              }
              

              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Sign Off's </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
              {
                this.state.fetchSignOffUsers&&this.state.fetchSignOffUsers.length<=5?<div className={styles.col_6}>
                <label>{this.state.SignOffList}</label>
              </div>:<Link onClick={this.showSignOffs}>All SignOff's</Link>
              } 
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
            </div>
            <hr />
            <div className={classnames(styles.row, "action-section")} hidden={this.state.btnsReadonly} >
              <div className={styles.col_7}>
                <label className={classnames(styles.divalign, "SignOff-sec")}>Action</label>
                {
                  this.state.StatusCheck == "SignOff Pending" ? <div className="check-sec"><Checkbox label="Sign Off" onChange={(e, option) => { this.setState({ chksignOffStatus: option }) }} /> </div>: <><ChoiceGroup defaultSelectedKey="A" options={this.options} onChange={(e, option) => { this.setState({ approveStatus: option.text }); }} required={true} /><div>
                    <div className={classnames(styles.col_6, "cmt-sec")}>
                      <label className={styles.divalign}>Comments</label>
                      <TextField multiline required={this.state.approveStatus == "Approve" ? false : true} value={this.state.comments} onChanged={newVal => {
                        newVal && newVal.length > 0
                          ? this.setState({
                            comments: newVal,
                            errorcomments: ""
                          })
                          : this.setState({
                            comments: newVal,
                          });
                      }} errorMessage={this.state.errorcomments} />
                    </div>
                  </div></>
                }

              </div>

            </div>
            {
              this.state.StatusCheck != "SignOff Pending" ? <>
                <div className={styles.textcenter}>
                  <span className={styles.buttonspace}>
                    <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this._alertClicked} />
                  </span>

                  <span className={styles.buttonspace} hidden={this.state.btnsReadonly}>
                    <PrimaryButton text="Submit" className={styles.btnSubmit} onClick={this.ApprovalSubmitForm} /></span>
                </div></> : <div className={styles.textcenter}>
                  <span className={styles.buttonspace}>
                    <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={(e) => this.setState({ topicValue: "", descriptionValue: "", fileDetails: [] })} />
                  </span>
                  <span className={styles.buttonspace} hidden={this.state.btnsReadonly}>
                    <PrimaryButton text="Submit" className={styles.btnSubmit} onClick={this.SignoffForm} /></span>
                </div>
            }

          </div> : this.queryMode == "view" && this.queryStringId ? <div className={styles.addTailGateRequest}>
            <div className="container">
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
                this.state.fetchApprovers&&this.state.fetchApprovers.length>0?<div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Approvers </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
            {
              this.state.fetchApprovers.length<=5?<div className={styles.col_6}>
              <label>{this.state.ApproversList}</label>
            </div>:<Link onClick={this.showApprover}>All Approvers</Link>
            }

              </div>:""
              }
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Sign Off's </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
              {
                this.state.fetchSignOffUsers&&this.state.fetchSignOffUsers.length<=5?<div className={styles.col_6}>
                <label>{this.state.SignOffList}</label>
              </div>:<Link onClick={this.showSignOffs}>All SignOff's</Link>
              } 
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
              <hr />
              <div className={classnames(styles.row, styles.nopaddingbottom,"history-contents")}>
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
            </div>
            <hr />
            { 
              <>
                <div className={styles.textcenter} style={{ paddingTop: "20px" }}>
                  <span className={styles.buttonspace}>
                    <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this._alertClicked} />
                  </span>
                </div></>
            }

          </div>  : this.queryMode == "editadmin" && this.queryStringId ? <div className={styles.addTailGateRequest}>
            <div className="container">
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
                this.state.fetchApprovers&&this.state.fetchApprovers.length>0?<div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Approvers </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
            {
              this.state.fetchApprovers.length<=5?<div className={styles.col_6}>
              <label>{this.state.ApproversList}</label>
            </div>:<Link onClick={this.showApprover}>All Approvers</Link>
            }

              </div>:""
              }
              <div className={classnames(styles.row, styles.nopaddingbottom)}>
                <div className={styles.col_3}>
                  <label className={classnames(styles.divalign, 'leftLbl')}>Sign Off's </label>
                </div>
                <div className={styles.col_1}>
                  :
            </div>
              {
                this.state.fetchSignOffUsers&&this.state.fetchSignOffUsers.length<=5?<div className={styles.col_6}>
                <label>{this.state.SignOffList}</label>
              </div>:<Link onClick={this.showSignOffs}>All SignOff's</Link>
              } 
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
            </div>
            <hr />
            <div className="history-contents">
              <div className={classnames(styles.row, "adminApproval")}>
                <div className={styles.col_6}>
                  <label className={styles.divalign}>Admin Actions</label>
                </div>
                <div className={classnames(styles.col_12, "tbl-margin")}> 
                 
                    {
                      
                     this.state.pendingAtStage=="Approval Stage"?
                     
                     this.approverstextArray.length>0&&this.approverstextArray.map((signOffItem,i) => {
                      var ItemDet="ApproverUser~"+signOffItem.Title+"~Approved";
                      var rowDet=this.state.ApproverAdminSummary;
                      if(rowDet)
                      {
                        var existIdx=rowDet.indexOf(ItemDet);
                        if(existIdx<0)
                        {
                          Dropoptions.push({key:signOffItem.ID,text:signOffItem.Title})
                        }
                  
                      }
                      else
                      {
                        Dropoptions.push({key:signOffItem.ID,text:signOffItem.Title});

                      }
                     
                     if(this.approverstextArray.length<=i+1)
                     {
                      return(      <><Dropdown
                        placeholder="Select an Approver"
                        label="Select an Approver"
                        options={Dropoptions}
                        styles={dropdownStyles}
                        onChange={(e,option)=>{this.setState({adminSelct:false,adminSelctApprover:option.text,adminApproverid:option.key.toString(),errorDropDown:""})}}
                        errorMessage={this.state.errorDropDown}
                        />

                        <div hidden={this.state.adminSelct}>
                        <ChoiceGroup defaultSelectedKey="A" options={this.options} onChange={(_e, option) => { this.setState({ approveStatus: option.text }); } } required={true} /><div>
                          <div className={classnames(styles.col_3, "cmt-sec")}>
                            <label className={styles.divalign}>Comments</label>
                            <TextField multiline required={this.state.approveStatus == "Approve" ? false : true} value={this.state.comments} onChanged={newVal => {
                              newVal && newVal.length > 0
                                ? this.setState({
                                  comments: newVal,
                                  errorcomments: ""
                                })
                                : this.setState({
                                  comments: newVal,
                                });
                            } } errorMessage={this.state.errorcomments} />
                          </div>
                        </div></div></>
                      )
                    
                     }
                  

                     })
                     
                     
                     :this.state.pendingAtStage=="SignOff Stage"?"":""
                    }
                 
                </div>
              </div>
            </div>
            {
             <div className={styles.textcenter}>
             <span className={styles.buttonspace}>
               <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this._alertClicked} />
             </span>
             <span className={styles.buttonspace} hidden={this.state.btnsReadonly}>
               <PrimaryButton text="Submit" className={styles.btnSubmit} onClick={this.AdminApprovalForm} /></span>
           </div>
            }

          </div>:"No records Found"
        }

      </div>
  

      </>

    );

  }
}
