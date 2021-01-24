import * as React from 'react';
import styles from './AddTailgateContractForm.module.scss';
import { IAddTailgateContractFormProps } from './IAddTailgateContractFormProps';
import { escape, times } from '@microsoft/sp-lodash-subset';
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
import "@pnp/sp/site-groups";
import { getGUID } from "@pnp/common";
import { Label } from 'office-ui-fabric-react/lib/Label';
import { ComboBox, IComboBoxOption, IComboBoxProps, IComboBox, SelectableOptionMenuItemType, flatten } from 'office-ui-fabric-react/lib/index';
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { DetailsList, DetailsRow, IDetailsListProps, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import {
  Dialog,
  DialogFooter,
  DialogType,
  IDialogStyles,
} from "office-ui-fabric-react/lib/Dialog";
import { IIconProps, IContextualMenuProps } from 'office-ui-fabric-react';
import { IconButton, PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { IComboBoxStyles, VirtualizedComboBox, Fabric } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IStackProps, Link, Stack } from 'office-ui-fabric-react';
import { IAddTailgateContractFormState } from './IAddTailgateContractFormState'
import classnames from 'classnames';
import swal from 'sweetalert';
import '../../../ExternalRef/style.css';
export default class AddTailgateContractForm extends React.Component<IAddTailgateContractFormProps, {}> {
  public state;
  public queryStringId:any;
  private contractorNumberMaxLength: number = 5;
  public dialogContentProps:any;
  listName: any = "TailgateTasksActivity";
  private _columns: IColumn[];
  options: IChoiceGroupOption[];
  public _classificationallItems: any[] = [];
  constructor(props: IAddTailgateContractFormProps) {
    super(props);
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });
    this.options = [
      { key: 'A', text: 'New Contractor' },
      { key: 'B', text: 'Re-Validation' }
    ];
    this.dialogContentProps = {
      title: '',
    };
    var currentURL = window.location.search.substring(1);
    var sURLVariables = currentURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) {
      var sParameterName = sURLVariables[i].split('=');
      if (sParameterName[0] == "IDCO") {
        this.state = { queryStringId: Number(sParameterName[1]) }

        this.queryStringId = Number(sParameterName[1]);
      }
   
    }
    this.state = {
      contractOptions: "New Contractor",
      classificationID:0,
      revalidationOptions: [],
      contractorNumber: "",
      errorcontractorNumber: "",
      contractorName: "",
      errorcontractorName: "",
      classification: [],
      errorclassification: "",
      validationPeriod: "",
      errorvalidationPeriod: "",
      filePickerResult: [],
      errorfileAttach: "",
      removedFiles:[],
      newFiles:[],
      TypeOfContract:"New Contractor",
      currentUserId:"",
      TypeOfContractId:"A",
      hideAlert:true,
      errorfilesizeMsg:"Attachment is required",
      HSEProcessType:"",
      PurchasingProcessType:"",
      FinanceProcessType:"",
      AttachmentType: [{key:1,text:"Safety policy & procedures"},{key:2,text:"Training"},{key:3,text:"Tailgate"},{key:4,text:"Violations"},{key:5,text:"General"}],
      errorAttachmentType: "",
      SelectedAttachementType:"",
      SelectedAttachementTypeID:0,

      
    }
    this.getCurrentUserDetails();
    this.getClassificationData();
  

    this.queryStringId!=undefined?this.getDraftData():"";
  }
  DeleteIcon: IIconProps = { iconName: 'Delete' };
  UploadIcon:IIconProps = { iconName: 'BulkUpload' };

  public async getCurrentUserDetails() {
    sp.web.currentUser.get().then((userId: any) => {
      this.setState({ currentUserId: userId.Id });
    });
  }
  public getClassificationData() {
    sp.web.lists.getByTitle('classification').items.get().then((olddatas: any) => {
      for (var i = 0; i < olddatas.length; i++) {
        var arritems2 = {
          key: olddatas[i]["Id"],
          text: olddatas[i]["Classification"]
        };
        this._classificationallItems.push(arritems2);
      }

      this.setState({ classification: this._classificationallItems });
      this.getProcessType();
    });
  }
  public async getProcessType() {
    sp.web.lists.getByTitle("ConfigList").items.select("*,UserName/Title,UserName/Id,UserName/EMail").expand('UserName').filter("Title eq 'HSE' or Title eq 'Purchasing' or Title eq 'Finance'").get().then((Ptypes: any) => {
      Ptypes.map((each)=>{
        if(each.Title == "HSE")
        {
          if(each.ProcessType=="User")
          this.setState({HSEProcessType:each.Title+"~User~"+each.UserName.Title+"~"+each.UserName.EMail})
          else if(each.ProcessType=="Group")
          this.setState({HSEProcessType:each.Title+"~Group~HSE_Approver"})

        }
        else if(each.Title == "Purchasing")
        {
          if(each.ProcessType=="User")
          this.setState({PurchasingProcessType:each.Title+"~User~"+each.UserName.Title+"~"+each.UserName.EMail})
          else if(each.ProcessType=="Group")
          this.setState({PurchasingProcessType:each.Title+"~Group~Purchasing"})
        }
        else if(each.Title == "Finance")
        {
          if(each.ProcessType=="User")
          this.setState({FinanceProcessType:each.Title+"~User~"+each.UserName.Title+"~"+each.UserName.EMail})
          else if(each.ProcessType=="Group")
          this.setState({FinanceProcessType:each.Title+"~Group~Finance"})
        }
      });
    });
  }

  private getDraftData() {
    sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).select("*,Title","Classification/Title","Classification/ID","AttachmentType").expand("Classification").get().then((items: any) => {
  
      this.setState({
        contractorNumber:items.ContractorNumber,
        contractorName: items.ContractorsName,
        classificationID: items.ClassificationId,
        validationPeriod: items.ReValidationPeriod,
        TypeOfContractId:items.TypeOfContract=="New Contractor"?"A":"B",
        HSEProcessType:items.HSEProcessType,
        PurchasingProcessType:items.PurchasingProcessType,
        FinanceProcessType:items.FinanceProcessType,
        SelectedAttachementTypeID:items.AttachmentType=="Safety policy & procedures"?1:items.AttachmentType=="Training"?2:items.AttachmentType=="Tailgate"?3:items.AttachmentType=="Violations"?4:items.AttachmentType=="General"?5:0,
        SelectedAttachementType:items.AttachmentType
      });
      var folderPath = "ContractorManagementDocuments/" + items.Id
      sp.web.getFolderByServerRelativeUrl(folderPath).files.select('*,ID').get().then((allFiles) => {

        var fetchFiles = [];
        allFiles.map((singleFile) => {
          fetchFiles.push({ filename: singleFile.Name, files: singleFile.ServerRelativeUrl })
        })
        this.setState({
          filePickerResult: fetchFiles
        });
      });
    }).catch(error => {

      this.setState({ errorcontractorNumber: "No Data availabe" })
    })
  }

  private getOldContractorNumber() {
    if(!this.state.contractorNumber)
    {
      this.setState({errorcontractorNumber: "Contract Number is required"});
      return false;
    }
    sp.web.lists.getByTitle("ContractorsManagement").items.select("*,Title", "Classification/Title", "Classification/ID","Created").expand("Classification").filter("ContractorNumber eq '" + this.state.contractorNumber + "' and FlowStatus eq 'Finance Approved'").orderBy("Created",false).get().then((items: any) => {

      this.setState({
        contractorName: items[0].ContractorsName,
        classificationID: items[0].ClassificationId,
        validationPeriod: items[0].ReValidationPeriod,
        currentUserId:items[0].ContractRequestorID,
        TypeOfContractId:items.TypeOfContract=="New Contractor"?"A":"B"
      });
      // var folderPath = "ContractorManagementDocuments/" + items[0].Id
      // sp.web.getFolderByServerRelativeUrl(folderPath).files.select('*,ID').get().then((allFiles) => {
     
      //   var fetchFiles = [];
      //   allFiles.map((singleFile) => {
      //     fetchFiles.push({ filename: singleFile.Name, files: singleFile.ServerRelativeUrl })
      //   })
      //   this.setState({
      //     filePickerResult: fetchFiles
      //   });
      // });
    }).catch(error => {

      this.setState({
        contractorName: "",
        classificationID: "",
        validationPeriod:"",
        filePickerResult:[]
      });
      this.setState({ errorcontractorNumber: "No Data availabe" })
    })
  }
  async EachfileUpload(allUploadFiles, result, newId,submitType) {
    await allUploadFiles.map((eachfileDetails, index) => {

      if (eachfileDetails.files["name"]) {
        result.folder.files.add(eachfileDetails.filename, eachfileDetails.files, true)
          .then((fresult) => {

            if (allUploadFiles.length <= index + 1) {
              this.setState({ filePickerResult: [] ,newFiles:[]});
              document.getElementById("loader-container").style.display= 'none';
              swal({
                title: 'Success',
                text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
                icon: 'success',              
              }).then(()=>{
                window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
              });


            }
          });
      }
    });
  }
  public toggleHideDialog = (event): void => {
    this.setState({hideAlert:true});
    window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
  }


  public cancelForm = (): void => {
    window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
  }
  public submitForm = (event): void => {
    var submitType = event.target.textContent == "Submit" ? "Submit" : "Draft";
    var d = new Date();
    var username=this.props.spcontext.pageContext.user.displayName;
    if(submitType=="Submit")
    {
      this.state.contractorNumber  ? "" : this.setState({ errorcontractorNumber: "Contractor Number is required" });
      this.state.contractorName  ? "" : this.setState({ errorcontractorName: "Contractor Name is required" });
      this.state.classificationID == 0 ? this.setState({ errorclassification: "Classification is required" }) : this.setState({ errorclassification: "" });
      this.state.filePickerResult.length > 0 ? "" : this.setState({ errorfileAttach: "Attachments are required" });
      this.state.validationPeriod  ? "" : this.setState({ errorvalidationPeriod: "Re-Validation Period is required" });
      this.state.SelectedAttachementTypeID==0  ?  this.setState({ errorAttachmentType: "Attachment type is required" }):this.setState({ errorAttachmentType: "" });
    }
    else
    {
      this.state.contractorNumber  ? "" : this.setState({ errorcontractorNumber: "Contractor Number is required" });
      this.setState({ errorcontractorName: "",errorclassification:"",errorvalidationPeriod: "", errorfileAttach: "" })
    }

    // this.state.contractorNumber.trim().length > 0 ? "" : this.setState({ errorcontractorNumber: "Contractor Number is required" });
    // this.state.contractorName.trim().length > 0 ? "" : this.setState({ errorcontractorName: "Contractor Name is required" });
    // this.state.classificationID == 0 ? this.setState({ errorclassification: "Classification is required" }) : this.setState({ errorclassification: "" });
    // this.state.filePickerResult.length > 0 ? "" : this.setState({ errorfileAttach: true });
    // this.state.validationPeriod.trim().length > 0 ? "" : this.setState({ errorvalidationPeriod: "Re-Validation Period is required" });
  

    
        if (this.queryStringId) {
          if((submitType=="Draft"&&this.state.contractorNumber)||(submitType=="Submit"&& (this.state.contractorNumber && this.state.contractorName && this.state.classificationID > 0 &&  this.state.validationPeriod && this.state.filePickerResult.length)))
          {
            document.getElementById("loader-container").style.display= 'flex'; 
            var year = d.getFullYear();
            var month = d.getMonth();
            var day = d.getDate();
           

        sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).update({
          Title: "Contractors Management",
          Status: event.target.textContent == "Submit" ? "Submit" : "Draft",
          HSEStatus: event.target.textContent == "Submit" ? "Pending" : "",
          ContractorsName: this.state.contractorName,
          ReValidationPeriod: this.state.validationPeriod,
          ContractorNumber: this.state.contractorNumber,
          ClassificationId: this.state.classificationID,
          ContractRequestorID:this.state.currentUserId.toString() ,
          FlowStatus: event.target.textContent == "Submit" ? "Request Submitted" : "Draft",   
          PendingAt:event.target.textContent == "Submit"? "HSE Stage" :"",
          AssignedDate:event.target.textContent == "Submit" ? new Date().toLocaleDateString()
          : ""  ,
          HSEProcessType:this.state.HSEProcessType,
          PurchasingProcessType:this.state.PurchasingProcessType,
          FinanceProcessType:this.state.FinanceProcessType,
          AutoTask:new Date(year + parseInt(this.state.validationPeriod), month, day).toLocaleDateString(),
          Level:"New Request~"+username+"~Created~"+new Date().toLocaleDateString()+"|",
          HSEApprovedCount:0,
          PurchasingApprovedCount:0,
          FinanceApprovedCount:0,
          AttachmentType:this.state.SelectedAttachementType.toString()
        }).then((disID: any) => {
         
          sp.web.getFolderByServerRelativeUrl("ContractorManagementDocuments").folders.add("ContractorManagementDocuments" + '/' + this.queryStringId).then(result => {
            var allUploadFiles = this.state.newFiles;
            var tobeRemove=this.state.removedFiles;
            if(tobeRemove.length)
            {
              tobeRemove.map((re)=>{
                sp.web.getFileByServerRelativeUrl(re.files).recycle().then(()=>{
      
                });
              });
            }
            if (allUploadFiles.length > 0) {
              this.EachfileUpload(allUploadFiles, result, this.queryStringId,submitType);
             

            }
            else {
              this.setState({ filePickerResult: [],newFiles:[] });
              document.getElementById("loader-container").style.display= 'none';
              swal({
                title: 'Success',
                text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
                icon: 'success',              
              }).then(()=>{
                window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
              });
             
            }
           // alert("Submitted Successfully..!");
          });
        });
      }
    }
      else
      {
        if((submitType=="Draft"&&this.state.contractorNumber)||(submitType=="Submit"&& (this.state.contractorNumber && this.state.contractorName && this.state.classificationID > 0 &&  this.state.validationPeriod && this.state.filePickerResult.length&&this.state.SelectedAttachementTypeID>0)))
        {
          document.getElementById("loader-container").style.display= 'flex'; 
     
          var year = d.getFullYear();
            var month = d.getMonth();
            var day = d.getDate();
              var attArray=[this.state.SelectedAttachementType]
        sp.web.lists.getByTitle("ContractorsManagement").items.add({
          Title: "Contractors Management",
          Status: event.target.textContent == "Submit" ? "Submit" : "Draft",
          HSEStatus: event.target.textContent == "Submit" ? "Pending" : "",
          ContractorsName: this.state.contractorName,
          ReValidationPeriod: this.state.validationPeriod,
          ContractorNumber: this.state.contractorNumber,
          ClassificationId: this.state.classificationID,
          TypeOfContract:this.state.contractOptions,
          ContractRequestorID:this.state.currentUserId.toString(),
          FlowStatus: event.target.textContent == "Submit" ? "Request Submitted" : "Draft",
          PendingAt:event.target.textContent == "Submit"? "HSE Stage" :"",
          AssignedDate:event.target.textContent == "Submit" ? new Date().toLocaleDateString()
          : "" ,
          HSEProcessType:this.state.HSEProcessType,
          PurchasingProcessType:this.state.PurchasingProcessType,
          FinanceProcessType:this.state.FinanceProcessType,  
          AutoTask:new Date(year + parseInt(this.state.validationPeriod), month, day).toLocaleDateString(),
          Level:"New Request~"+username+"~Created~"+new Date().toLocaleDateString()+"|",
          HSEApprovedCount:0,
          PurchasingApprovedCount:0,
          FinanceApprovedCount:0,
          AttachmentType:this.state.SelectedAttachementType.toString()
          // AttachmentType:{"results": attArray }
        })
          .then((disID: any) => {
          
            sp.web.getFolderByServerRelativeUrl("ContractorManagementDocuments").folders.add("ContractorManagementDocuments" + '/' + disID.data.Id).then(result => {
              var allUploadFiles = this.state.filePickerResult;
              if (allUploadFiles.length > 0) {
                this.EachfileUpload(allUploadFiles, result, disID.data.Id,submitType);
              }
              else {
                this.setState({ filePickerResult: [] });
                document.getElementById("loader-container").style.display= 'none';
                swal({
                  title: 'Success',
                  text: submitType == "Submit" ? "Request Submitted Successfully..!!" : "Request Saved successfully..!!",
                  icon: 'success',              
                }).then(()=>{
                  window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
                });
                

              }
           
             // alert("Submitted Successfully..!");
            });
          });
      }

    }
    

  }

  public fileUploadCallback = (e) => {
    if(!this.queryStringId)
    {
      var files = e.target.files;
  //   files[0].size<=5000000?files=e.target.files: files="";this.setState({errorfileAttach:"File should not be more than 5 MB"})
      if (files && files.length > 0 && files[0].size<=5000000) {
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
      else if(files && files.length > 0 && files[0].size>=5000000){
        this.setState({ filePickerResult: [], errorfileAttach: "Attachments file should not be more than 5 MB" });
        e.target.value = null;
      }     
      else {
        this.setState({ filePickerResult: [], errorfileAttach: "Qualification documents are required" });
        e.target.value = null;
      }
    }
    else
    {
      var files = e.target.files;
      if (files && files.length > 0) {
        var allfiles = [];
        var oldallfiles=[];
        for (let i = 0; i < files.length; i++) {
          if (this.state.newFiles)
            allfiles = this.state.newFiles;
            if (this.state.filePickerResult)
            oldallfiles = this.state.filePickerResult;
          var sepArray = allfiles.filter((eleFile) => { return eleFile.filename == files[i].name })
          if (sepArray.length <= 0) {
            allfiles.push({ filename: files[i].name, files:  files[i] });
            oldallfiles.push({ filename: files[i].name, files:  files[i]})
          }
  
          if (files.length <= i + 1)
            this.setState({ newFiles: allfiles,filePickerResult:oldallfiles });
        }
        e.target.value = null;
      } else {
        this.setState({ newFiles: [] , errorfileAttach: "Qualification documents are required" });
        e.target.value = null;
      }
    }
    
  }
  public removeDoc = (e) => {
    var targetelement;
    if(!this.queryStringId)
    {
       targetelement = e.currentTarget.id;
      var filesArray = this.state.filePickerResult;
      filesArray = filesArray.filter((key, index) => { return index != targetelement; });
     
      this.setState({ filePickerResult: filesArray });
    }
    else
    {
       targetelement = parseInt(e.currentTarget.id);
      var filesArray = this.state.filePickerResult;
      var removedFiles=this.state.removedFiles;
      removedFiles=filesArray.filter((key, index) => {
     return index == targetelement
           });
  
      filesArray = filesArray.filter((key, index) => {
      return index!=targetelement
           });
      // this.file = null;
      this.setState({ filePickerResult: filesArray ,removedFiles:removedFiles });
    }
  
  }
  public render(): React.ReactElement<IAddTailgateContractFormProps> {
   
    return (
      <div className={styles.addTailgateContractForm}>
        <div id="loader-container" style={{display:"none"}}>
            <div className="loader">
              <span></span>
              <span></span>
              <span></span>
              <span></span>
              <span></span>
            </div>
          </div>
        <h2>Contractor Management</h2>
        <div className={styles.container}>
          <div className={styles.row}>         
            <div className={styles.col_8}>              
              <ChoiceGroup selectedKey={this.state.TypeOfContractId} options={this.options} onChange={(e, option) => {
              this.setState({
                  contractorName: "",
                  classification: this._classificationallItems,
                  errorcontractorNumber: "Contractor Number is required",
                  classificationID: "",
                  validationPeriod: "",
                  filePickerResult: [],
                  contractOptions: option.text,
                  contractorNumber:"",
                  TypeOfContractId:option.key
                });
              }} required={true} />
              </div></div> <div className="contract-req">  <div className={styles.row}>
              <div className={styles.col_2}>
              <Label className={styles.divalign} required>Contractor Number</Label>         
         </div>
            <div className={styles.col_6}>
              <TextField 
                value={this.state.contractorNumber}
                onChanged={newVal => {
                  var letters = /^[0-9]+$/;
                  var numberCheck;
                  numberCheck = newVal.match(letters) ? false : true;
                  if (newVal && newVal.length > 0) {
                    if (numberCheck == false && newVal.length > this.contractorNumberMaxLength) {
                      this.setState({ errorcontractorNumber: "Contract Number should be accept 5 digits" });
                    }
                    else if (numberCheck) {
                      this.setState({ errorcontractorNumber: "Contract Number should be accept letters" });
                    }
                    else {
                      this.setState({ contractorNumber: newVal, errorcontractorNumber: "" });
                    }
                  }
                  // else {
                  //   this.setState({ contractorNumber: newVal, errorcontractorNumber: "Contract Number is required" });
                  // }
                
                }
              }
                errorMessage={this.state.errorcontractorNumber}></TextField>
                {
             this.state.contractOptions=="Re-Validation" ?    <div className={"contBtn"}>
              <PrimaryButton className={styles.btnDraft} text="Get contractor details" onClick={(e)=>this.getOldContractorNumber.call(this,e)} />
            </div>:""
                }
              
                </div><div>
            </div>
          </div>
          <div className={styles.row}>
          <div className={styles.col_2}>
          <Label className={styles.divalign} required>Contractor Name</Label>         
          </div>
            <div className={styles.col_6}>
              <TextField 
                value={this.state.contractorName}
                onChanged={newVal => {
                  newVal && newVal.length > 60
                    ? this.setState({
                      contractorName: newVal,
                      errorcontractorName: "Contractor Name should not be more than 60 Characters"
                    })
                    : this.setState({
                      contractorName: newVal,
                      errorcontractorName:
                        ""
                    });
                }}  errorMessage={this.state.errorcontractorName}></TextField> 
            </div>
          </div>
          <div className={styles.row}>
          <div className={styles.col_2}>
              <Label className={styles.divalign} required>Classification</Label>         
         </div>
            <div className={styles.col_6}>
              <VirtualizedComboBox
                placeholder={"Select an option"}            
                onChange={(e, option) => { this.setState({ classificationID: Number(option.key) }) }}
                selectedKey={this.state.classificationID}                
                allowFreeform
                autoComplete="on"
                options={this.state.classification}
                dropdownMaxWidth={200}
                useComboBoxAsMenuWidth
                required={true}
                errorMessage={this.state.errorclassification}
              />
           
            </div>
          </div>

          <div className={styles.row}>
          <div className={styles.col_2}>
              <Label className={styles.divalign} required>Re-Validation Period</Label>         
         </div>
            <div className={styles.col_6}>
              <TextField 
                value={this.state.validationPeriod}
                onChanged={newVal => {
                  var letters = /^[0-9]+$/;
                                var numberCheck;
                                numberCheck = newVal.match(letters) ? false : true;
                                if (newVal && newVal.length > 0) {
                                  if (numberCheck == false && newVal.length > this.contractorNumberMaxLength) {
                                    this.setState({ errorvalidationPeriod: "Re-Validation Period should be accept 5 numbers" });
                                  }
                                  else if (numberCheck) {
                                    this.setState({ errorvalidationPeriod: "Re-Validation Period should be accept letters" });
                                  }
                                  else {
                                    this.setState({ validationPeriod: newVal, errorvalidationPeriod: "" });
                                  }
                                }
                                else {
                                  this.setState({ validationPeriod: newVal, errorvalidationPeriod: "Re-Validation Period is required" });
                                }
                }
                } errorMessage={this.state.errorvalidationPeriod}></TextField> 
            </div>
          </div>

          <div className={styles.row}>
          <div className={styles.col_2}>
              <Label className={styles.divalign} required>Upload Qualification Form</Label>         
         </div>
            <div className={styles.col_2}>
              <div>   
              <div className="custom-upload">
                       
                       <input  id="fileCus" disabled={this.state.viewMode} type="file" multiple accept=".xlsx,.xls,.doc, image/*, .docx,.ppt, .pptx,.txt,.pdf" onChange={this.fileUploadCallback} />
                       <label htmlFor="fileCus" > <IconButton className={"file-upload-icon"} iconProps={this.UploadIcon}>  
                 </IconButton>Choose File</label>
                     </div>            
                {/* <input type="file" accept=".xlsx,.xls,.doc, image/*, .docx,.pdf" onChange={this.fileUploadCallback}
                /> */}
                {
                  this.state.filePickerResult ?
                    this.state.filePickerResult.map((filedet, index) => {
                      return (
                        <div className={styles.attach}>
                          <label style={{ color: "#333" }}>{filedet.filename}</label>
                          <IconButton className={styles.btntransparent} iconProps={this.DeleteIcon} onClick={this.removeDoc.bind(this)} id={index.toString()}>
                          </IconButton><br></br>
                        </div>
                      );

                    }) : ""
                }
             
                {this.state.errorfileAttach ? <Label className={styles.pickerlabelErrormsg}>{this.state.errorfileAttach}</Label> : ""}
              </div>
            </div>

            
            <div className={styles.col_2}>
              <VirtualizedComboBox
                placeholder={"Select an option"}            
                onChange={(e, option) => { this.setState({ SelectedAttachementType:option.text,SelectedAttachementTypeID:option.key}) }}
                selectedKey={this.state.SelectedAttachementTypeID}                
                allowFreeform
                autoComplete="on"
                options={this.state.AttachmentType}
                dropdownMaxWidth={200}
                useComboBoxAsMenuWidth
                required={true}
                errorMessage={this.state.errorAttachmentType}
              />
           
            </div>
            <div className={styles.col_2}>
         </div>
          </div> 


          </div>   
          <div className={classnames(styles.row, "btnscontract")}>
            <div className={classnames(styles.col_3, styles.btnCancel)}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div>
            <div className={styles.col_3}>
              <PrimaryButton className={styles.btnDraft} text="Save as Draft" onClick={this.submitForm} />
            </div>
            <div className={styles.col_3}>
              <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitForm} />
            </div>
          </div>
        </div>     
         <div>   <Dialog
        hidden={this.state.hideAlert}      
        dialogContentProps={this.dialogContentProps}      
      >
        <DialogFooter>
          <PrimaryButton onClick={this.toggleHideDialog} text="Ok" />      
        </DialogFooter></Dialog>
        </div>
      </div>
    
    );
  }
}
