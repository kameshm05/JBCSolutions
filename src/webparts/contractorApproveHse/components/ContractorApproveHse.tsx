import * as React from 'react';
import styles from './ContractorApproveHse.module.scss';
import { IContractorApproveHseProps } from './IContractorApproveHseProps';
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
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import {
  Dialog,
  DialogFooter,
  DialogType,
  IDialogStyles,
} from "office-ui-fabric-react/lib/Dialog";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IStackProps, Link, Stack } from 'office-ui-fabric-react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IIconProps, IContextualMenuProps } from 'office-ui-fabric-react';
import { IconButton, PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { IComboBoxStyles, VirtualizedComboBox, Fabric } from 'office-ui-fabric-react';
import { IContractorApproveHseState } from './IContractorApproveHseState';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

import classnames from 'classnames';
import swal from 'sweetalert';
import '../../../ExternalRef/style.css';import '../../../ExternalRef/style.css';
export default class ContractorApproveHse extends React.Component<IContractorApproveHseProps, {}> {
  public state;
  public queryStringId: any;
  public queryMode:any;
  options: IChoiceGroupOption[];
  purchaseoptions: IChoiceGroupOption[];
  public HSEtextArray=[];
  public PurchasingtextArray=[];
  public FinancetextArray=[];
  constructor(props: IContractorApproveHseProps) {
    super(props);
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      },
    });
    this.options = [
      { key: 'A', text: 'Approve' },
      { key: 'B', text: 'Return' }
    ];
    this.purchaseoptions = [{ key: 'A', text: 'Approve' },
    { key: 'B', text: 'In-Progress' }]

    var currentURL = window.location.search.substring(1);
    var sURLVariables = currentURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) {
      var sParameterName = sURLVariables[i].split('=');
      if (sParameterName[0] == "IDCO") {
        this.state = { queryStringId: Number(sParameterName[1]) }

        this.queryStringId = Number(sParameterName[1]);
      }
      else if(sParameterName[0] == "CMode")
      {
        this.queryMode=sParameterName[1].toLowerCase();
      }
    }

    this.state = {
      approveOptions: "Approve",
      approvepurchaseOptions: "Approve",
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
      currentUserId: "",
      isHSEGroupApprover: false,
      isPurchasingGroupApprover: false,
      isFinanceGroupApprover: false,
      hideHseApprove: false,
      hidePurchasingApprove: false,
      hideFinanceApprove: false,
      ispurchasingApprove: false,
      isPurchasingStatus: "",
      isPurchasingLevel: "",
      isPurchasingDate: "",
      axNumber: "",
      erroraxNumber: "",
      isnotApprover:false,
      LevelSummary:"",
      HSEStatus:"",
      PurchasingStatus:"",
      FinanceStatus:"",
      endUserHide:false,
      hidereturnBox:true,
      returncomments:"",
      errorreturncomments:"",
      hideAlert:true,
      HSEProcessType:"",
      PurchasingProcessType:"",
      FinanceProcessType:"",
      HSENeededCount:0,
      HSEApprovedCount:0,
      PurchasingNeededCount:0,
      PurchasingApprovedCount:0,
      FinanceNeededCount:0,
      FinanceApprovedCount:0,
      TypeOfContract:"",
      adminSelct:true,
      adminSelctApprover:"",
      errorDropDown:"",
      pendingAtStage:"",
      adminApproverid:"",
      AttachmentType:""
    }

    // this.getApproveData();
    this.getApproverDetails();

  }
  public getgroupUserCount=(grpName)=>{
    sp.web.siteGroups.getByName(grpName).users.get().then((result)=> {
      if(grpName=="HSE_Approver")
      {
        result.map((user)=>{
          this.HSEtextArray.push({Title:user.Title,ID:user.Id});
        });
        var count=result.length;
        this.setState({HSENeededCount:count})
      }
      else if(grpName=="Purchasing")
      {
        result.map((user)=>{
          this.PurchasingtextArray.push({Title:user.Title,ID:user.Id});
        });
        var count=result.length;
        this.setState({PurchasingNeededCount:count})
      }
      else if(grpName=="Finance")
      {
        result.map((user)=>{
          this.FinancetextArray.push({Title:user.Title,ID:user.Id});
        })
        var count=result.length;
        this.setState({FinanceNeededCount:count})
      }

    });
  }

  public async getApproverDetails() {
    this.setState({  approveOptions: "Approve",approvepurchaseOptions: "Approve"});
    sp.web.currentUser.get().then((userId: any) => {
      this.setState({ currentUserId: userId.Id })
    });
    let groups = await sp.web.currentUser.groups();
    for (var i = 0; i < groups.length; i++) {
      if (groups[i].LoginName == "HSE_Approver") {
        this.setState({ isHSEGroupApprover: true,isnotApprover:false});
      }
      else if (groups[i].LoginName == "Purchasing") {
        this.setState({ isPurchasingGroupApprover: true,isnotApprover:false });
      }
      else if (groups[i].LoginName == "Finance") {
        this.setState({ isFinanceGroupApprover: true,isnotApprover:false });
      }
      



    }
    if(this.queryMode=="view")
    this.setState({isnotApprover:true});
    if(groups.length<=i+1)
    this.getApproveData();

    console.log("isHSEGroupApprover" + this.state.isHSEGroupApprover);
    console.log("Purchasing" + this.state.isPurchasingGroupApprover);
    console.log("Finance" + this.state.isFinanceGroupApprover);
  }
  public getApproveData() {
    sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).select("*,Title", "Classification/Classification", "Classification/ID").expand("Classification").get().then((items: any) => {
      console.log("Items" + items);
      this.setState({
        contractorNumber: items.ContractorNumber,
        contractorName: items.ContractorsName,
        classificationID: items.Classification.Classification,
        validationPeriod: items.ReValidationPeriod,
        hideHseApprove: items.HSEStatus == "Approved" ||  items.HSEStatus == "Returned"? false : true,
        hidePurchasingApprove: items.PurchasingStatus == "Pending"||  items.HSEStatus == "In-Progress" ? true : false,
        LevelSummary:items.Level,
        HSEStatus:items.HSEStatus,
        PurchasingStatus:items.PurchasingStatus,
        FinanceStatus:items.FinanceStatus,
        axNumber:items.AXNumber,
        HSEApprovedCount:items.HSEApprovedCount,
        TypeOfContract:items.TypeOfContract,
        PurchasingApprovedCount:items.PurchasingApprovedCount,
        FinanceApprovedCount:items.FinanceApprovedCount,
        pendingAtStage:items.PendingAt,
        AttachmentType:items.AttachmentType
      });
    if(items.HSEStatus=="Pending")
    {
      if(items.HSEProcessType.indexOf('HSE~User')>=0)
      {
        this.setState({HSEProcessType:"User"});
        this.setState({HSENeededCount:1,isHSEGroupApprover:true});
        var splitval=items.HSEProcessType.split('~')
        this.HSEtextArray.push({Title:splitval[2],ID:1});
      }
      else if(items.HSEProcessType.indexOf('HSE~Group')>=0)
      {
        this.setState({HSEProcessType:"Group"});
        this.getgroupUserCount('HSE_Approver');
      }
    }
    else if(items.PurchasingStatus=="Pending")
    {
      if(items.PurchasingProcessType.indexOf('Purchasing~User')>=0)
      {
        this.setState({PurchasingProcessType:"User"});
        this.setState({PurchasingNeededCount:1});
        var splitval=items.PurchasingProcessType.split('~')
        this.PurchasingtextArray.push({Title:splitval[2],ID:1});
      }
      else if(items.PurchasingProcessType.indexOf('Purchasing~Group')>=0)
      {
        this.setState({PurchasingProcessType:"Group"});
        this.getgroupUserCount('Purchasing');
      }
    }
    else if(items.FinanceStatus=="Pending")
    {
      if(items.FinanceProcessType.indexOf('Finance~User')>=0)
      {      
        this.setState({FinanceProcessType:"User"});
        this.setState({FinanceNeededCount:1});
        var splitval=items.FinanceProcessType.split('~')
        this.FinancetextArray.push({Title:splitval[2],ID:1});
      }
      else if(items.FinanceProcessType.indexOf('Finance~Group')>=0)
      {
        this.setState({FinanceProcessType:"Group"});
        this.getgroupUserCount('Finance');
      }
    }

  if((items.PurchasingStatus=="Pending"||items.PurchasingStatus=="In-Progress") && this.state.isPurchasingGroupApprover)
  {
    this.setState({ hideHseApprove:false,hideFinanceApprove:false, hidePurchasingApprove:true})
  }
  if((items.PurchasingStatus=="Approved"||items.PurchasingStatus=="In-Progress"||items.FinanceStatus=="In-Progress"||items.FinanceStatus=="Pending")&&this.state.isFinanceGroupApprover&&items.FinanceStatus!="Approved")
  {
    this.setState({ hideHseApprove:false,hidePurchasingApprove:false, hideFinanceApprove:true})
  }
 
  var folderPath = "ContractorManagementDocuments/" + this.queryStringId
  sp.web.getFolderByServerRelativeUrl(folderPath).files.select('*,ID').get().then((allFiles) => {
    console.log(allFiles);
    var fetchFiles = [];
    allFiles.map((singleFile) => {
      fetchFiles.push({ filename: singleFile.Name, files: singleFile.ServerRelativeUrl })
    })
    this.setState({
      filePickerResult: fetchFiles
    });
  });


    }).catch(error => {
      console.log(error)
      this.setState({ errorcontractorNumber: "No Data availabe" })
    })
  }
  public cancelForm = (event): void => {
    window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";

  }
  public submitFormHSE = (event): any => {
    var typeofSubmit=event.currentTarget.id;
    if(typeofSubmit=="adminsubmit")
    {
      if(!this.state.adminSelctApprover)
      {
        this.setState({errorDropDown:"Please select approver"});
        return false
      }
    }
    var date = new Date().toLocaleString();
    var obj={};
    var historySummary;
    var neededCount=this.state.HSENeededCount;
    var ApprovedCount=this.state.HSEApprovedCount+1;
    var currentUserName="";
    if(this.state.LevelSummary)
     historySummary=this.state.LevelSummary;
    else
     historySummary="";
    if(!this.state.adminSelct)
     currentUserName=this.state.adminSelctApprover;
    else
     currentUserName=this.props.spcontext.pageContext.user.displayName;
    var summary= this.state.approveOptions=="Approve"?"Approved":"Returned";
 
     obj={ 
     HSEStatus: this.state.approveOptions=="Approve" && neededCount==ApprovedCount?"Approved":this.state.approveOptions=="Return"?"Returned":this.state.approveOptions=="Approve"?"Pending":"",
     ApprovedHSEDate: date,
     PurchasingStatus:this.state.approveOptions=="Approve" && neededCount==ApprovedCount?"Pending":"",
     Level:historySummary+"HSE~"+currentUserName+"~"+summary+"~"+date+"|",
     FlowStatus: this.state.approveOptions=="Approve" && neededCount==ApprovedCount?"HSE Approved":this.state.approveOptions=="Return"?"HSE Returned":"",
     PendingAt:this.state.approveOptions=="Approve" && neededCount==ApprovedCount? "Purchasing Stage" :this.state.approveOptions=="Return"?"Returned Stage":"HSE Stage",
     AssignedDate:new Date().toLocaleDateString(),
     HSEApprovedCount:ApprovedCount
    }

     if( this.state.approveOptions=="Approve")
     {
      sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).update(
        obj
      ).then(_success => {
        swal({
          title: 'Success',
          text:  "Request updated successfully..!!",
          icon: 'success',              
        }).then(()=>{
          window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
        });
      })
     }
     else if(  this.state.approveOptions=="Return" && this.state.returncomments.length>0){
      sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).update(obj).then(_success => {
        swal({
          title: 'Success',
          text:  "Request updated successfully..!!",
          icon: 'success',              
        }).then(()=>{
          window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
        });
      })
    }
    else{
      this.setState({errorreturncomments:"Comments is required" });
    }
  }

  public submitFormPurchasing = (event): any => {
    var typeofSubmit=event.currentTarget.id;
    if(typeofSubmit=="adminsubmit")
    {
      if(!this.state.adminSelctApprover)
      {
        this.setState({errorDropDown:"Please select approver"});
        return false
      }
    }

    var date = new Date().toLocaleString();
    var summary= this.state.approvepurchaseOptions=="Approve"?"Approved":"In-Progress"
    var neededCount=this.state.PurchasingNeededCount;
    var ApprovedCount=this.state.PurchasingApprovedCount+1;
    var historySummary=this.state.LevelSummary;
    var currentUserName="";

    if(!this.state.adminSelct)
    currentUserName=this.state.adminSelctApprover;
   else
    currentUserName=this.props.spcontext.pageContext.user.displayName;

    sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).update({
      PurchasingStatus: (this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Approved":"Pending",
      ApprovedPurchasingDate: date,
      Level:historySummary+"Purchasing~"+currentUserName+"~"+summary+"~"+date+"|",
      FinanceStatus:(this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Pending":"",
      FlowStatus: (this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Purchasing Approved":"",
      PendingAt:(this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Finance Stage":"Purchasing Stage",
      AssignedDate:new Date().toLocaleDateString() ,
      PurchasingApprovedCount:ApprovedCount
    }).then(_success => {
      swal({
        title: 'Success',
        text:  "Request updated successfully..!!",
        icon: 'success',              
      }).then(()=>{
        window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
      });
    });
  }
  public submitFormFinance = (event): any => {
    var typeofSubmit=event.currentTarget.id;
    if(typeofSubmit=="adminsubmit")
    {
      if(!this.state.adminSelctApprover)
      {
        this.setState({errorDropDown:"Please select approver"});
        return false
      }
    }

    var date = new Date().toLocaleString();
    var summary= this.state.approvepurchaseOptions=="Approve"?"Approved":"In-Progress";
    var neededCount=this.state.FinanceNeededCount;
    var ApprovedCount=this.state.FinanceApprovedCount+1;
    var historySummary=this.state.LevelSummary;
    var currentUserName="";
    
    if(!this.state.adminSelct)
    currentUserName=this.state.adminSelctApprover;
   else
    currentUserName=this.props.spcontext.pageContext.user.displayName;

      if(this.state.approvepurchaseOptions=="Approve"&&!this.state.axNumber)
      {
        this.setState({erroraxNumber:"AX Number is required"});
      }
      else
      {
        this.setState({erroraxNumber:""});
      }
    var historySummary=this.state.LevelSummary;
    if((this.state.axNumber&&this.state.approvepurchaseOptions=="Approve")||this.state.approvepurchaseOptions=="In-Progress"){
    sp.web.lists.getByTitle("ContractorsManagement").items.getById(this.queryStringId).update({
      FinanceStatus: (this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Approved":"Pending",
      ApprovedFinanceDate: date,
      Level: historySummary+"Finance~"+currentUserName+"~"+summary+"~"+date+"|",
      AXNumber: this.state.axNumber,
      FlowStatus: (this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Finance Approved":"",
      PendingAt:(this.state.approvepurchaseOptions=="Approve"|| this.state.approvepurchaseOptions=="In-Progress") && neededCount==ApprovedCount?"Completed Stage":"Finance Stage",
      AssignedDate:new Date().toLocaleDateString(),
      FinanceApprovedCount:ApprovedCount
    }).then(_success => {
      swal({
        title: 'Success',
        text:  "Request updated successfully..!!",
        icon: 'success',              
      }).then(()=>{
        window.location.href=this.props.siteUrl+"/SitePages/TailGateRequestDashBoard.aspx";
      });

    });
  }else{
    this.setState({erroraxNumber:"AX Number is required"})
  }
  }
  public render(): React.ReactElement<IContractorApproveHseProps> {
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 300 },
    };
    const dialogContentProps = {
      subText: 'Contract Save as draft or Submitted Successfully..!!',
    };
    const Dropoptions: IDropdownOption[] = [

    ];
    var StatusSummary=[];
    if(this.state.LevelSummary)
     StatusSummary=this.state.LevelSummary.split('|');
    return (
      <div className={styles.contractorApproveHse}>
       {
         this.state.hideHseApprove && this.state.isHSEGroupApprover?<h2>Contractor Management - HSE</h2>:this.state.hidePurchasingApprove && this.state.isPurchasingGroupApprover?<h2>Contractor Management - Purchasing</h2>: this.state.hideFinanceApprove && this.state.isFinanceGroupApprover?<h2>Contractor Management - Finance</h2>:""
       } 
        <div className={styles.container}>
          <h3>Info Contractor</h3>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Status</label>
            </div>
            <div className={styles.col_6}>
              <label>{this.state.TypeOfContract}</label>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Contractor Number</label>
            </div>
            <div className={styles.col_6}> 
              <label>{this.state.contractorNumber}</label>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Contractor Name</label>
            </div>
            <div className={styles.col_6}>
              <label>{this.state.contractorName}</label>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Classification</label>
            </div>
            <div className={styles.col_6}>
              <label>{this.state.classificationID}</label>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Re-Validation Period/Years</label>
            </div>
            <div className={styles.col_6}>
              <label>{this.state.validationPeriod}</label>
            </div>
          </div>
          {
            this.state.axNumber?<div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>AX Number</label>
            </div>
            <div className={styles.col_6}>
              <label>{this.state.axNumber}</label>
            </div>
          </div>:""
          }
          
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Upload Qualification Form</label>
            </div>
            <div className={styles.col_6}>
              {
                this.state.filePickerResult.map((filedet) => {
                  return (
                    <div>
                      <Link href={filedet.files}>{filedet.filename}</Link>

                    </div>
                  )
                })
              }
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Type</label>
            </div>
            <div className={styles.col_6}>
              <label>{this.state.AttachmentType}</label>
            </div>
          </div>
          <hr/>
          {/* { HSE State Level        } */}
        {
         ! this.state.isnotApprover?
         <div>
           
            {
               this.queryMode=="editadmin" ?   
               <div >
                {
                  StatusSummary&&StatusSummary.length>0? <div>
                  <div className={styles.row}>
              <h3>History Contractor</h3>
                <div className={styles.col_12}>
                  <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead ><tbody>
                  {
                    StatusSummary.map((rowDet)=>{
                      if(rowDet)
                      {
                       rowDet=rowDet.split('~');
                  return(<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>)
                    
                    }})
                  }
                  
                  </tbody></table>
                </div>
              </div><hr/></div>:""
                }
                <div className="history-contents">
              <div className={classnames(styles.row, "adminApproval")}>
                <div className={styles.col_6}>
                  <label className={styles.divalign}>Admin Actions</label>
                </div>
                <div className={classnames(styles.col_12, "tbl-margin")}> 
                 
                    {
                      
                     this.state.pendingAtStage=="HSE Stage"?
                     
                     this.HSEtextArray.length>0&&this.HSEtextArray.map((HSEItem,i) => {
                      var ItemDet="HSE~"+HSEItem.Title+"~Approved";
                      var rowDet=this.state.LevelSummary;
                      if(rowDet)
                      {
                        var existIdx=rowDet.indexOf(ItemDet);
                        if(existIdx<0)
                        {
                          Dropoptions.push({key:HSEItem.ID,text:HSEItem.Title})
                        }
                  
                      }
                      else
                      {
                        Dropoptions.push({key:HSEItem.ID,text:HSEItem.Title});

                      }
                     
                     if(this.HSEtextArray.length<=i+1)
                     {
                      return(      <><Dropdown
                        placeholder="Select an user"
                        label="Select an user"
                        options={Dropoptions}
                        styles={dropdownStyles}
                        onChange={(e,option)=>{this.setState({adminSelct:false,adminSelctApprover:option.text,adminApproverid:option.key.toString(),errorDropDown:""})}}
                        errorMessage={this.state.errorDropDown}
                        />

                        <div hidden={this.state.adminSelct}>
                        <ChoiceGroup defaultSelectedKey="A" options={this.options} onChange={(_e, option) => {option.text=="Approve"? this.setState({ approveOptions: option.text,hidereturnBox:true }):this.setState({ approveOptions: option.text,hidereturnBox:false }) } } required={true} />    <div className={styles.col_6} hidden={this.state.hidereturnBox}>
            <TextField label="Comments" required={true}
                value={this.state.returncomments}
                onChanged={newVal => {
                  newVal && newVal.length > 0
                    ? this.setState({
                      returncomments: newVal,
                      errorreturncomments: ""
                    })
                    : this.setState({
                      returncomments: newVal,
                      errorreturncomments: "Comments is required"
                    })
                }}
                errorMessage={this.state.errorreturncomments}></TextField>
            </div></div></>
                      )
                    
                     }
                  

                     })
                     
                     
                     :this.state.pendingAtStage=="Purchasing Stage"?
                     this.PurchasingtextArray.length>0&&this.PurchasingtextArray.map((PItem,i) => {
                      var ItemDet="Purchasing~"+PItem.Title+"~Approved";
                      var ItemDet1="Purchasing~"+PItem.Title+"~In-Progress";
                      var rowDet=this.state.LevelSummary;
                      if(rowDet)
                      {
                        var existIdx=rowDet.indexOf(ItemDet);
                        var existIdx1=rowDet.indexOf(ItemDet1);
                        if(existIdx<0&&existIdx1<0)
                        {
                          Dropoptions.push({key:PItem.ID,text:PItem.Title})
                        }
                  
                      }
                      else
                      {
                        Dropoptions.push({key:PItem.ID,text:PItem.Title});

                      }
                     
                     if(this.PurchasingtextArray.length<=i+1)
                     {
                      return(      <><Dropdown
                        placeholder="Select an user"
                        label="Select an user"
                        options={Dropoptions}
                        styles={dropdownStyles}
                        onChange={(e,option)=>{this.setState({adminSelct:false,adminSelctApprover:option.text,adminApproverid:option.key.toString(),errorDropDown:""})}}
                        errorMessage={this.state.errorDropDown}
                        />

                        <div hidden={this.state.adminSelct}>
                        <ChoiceGroup defaultSelectedKey="A" options={this.purchaseoptions} onChange={(_e, option) => { this.setState({ approvepurchaseOptions: option.text }); } } required={true} /></div></>
                      )
                    
                     }
                  

                     })
                     :this.state.pendingAtStage=="Finance Stage"?
                     this.FinancetextArray.length>0&&this.FinancetextArray.map((PItem,i) => {
                      var ItemDet="Finance~"+PItem.Title+"~Approved";
                      var ItemDet1="Finance~"+PItem.Title+"~In-Progress";
                      var rowDet=this.state.LevelSummary;
                      if(rowDet)
                      {
                        var existIdx=rowDet.indexOf(ItemDet);
                        var existIdx1=rowDet.indexOf(ItemDet1);
                        if(existIdx<0&&existIdx1<0)
                        {
                          Dropoptions.push({key:PItem.ID,text:PItem.Title})
                        }
                  
                      }
                      else
                      {
                        Dropoptions.push({key:PItem.ID,text:PItem.Title});

                      }
                     
                     if(this.FinancetextArray.length<=i+1)
                     {
                      return(      <><Dropdown
                        placeholder="Select an user"
                        label="Select an user"
                        options={Dropoptions}
                        styles={dropdownStyles}
                        onChange={(e,option)=>{this.setState({adminSelct:false,adminSelctApprover:option.text,adminApproverid:option.key.toString(),errorDropDown:""})}}
                        errorMessage={this.state.errorDropDown}
                        />

                        <div hidden={this.state.adminSelct}>
                        <ChoiceGroup defaultSelectedKey="A" options={this.purchaseoptions} onChange={(_e, option) => { this.setState({ approvepurchaseOptions: option.text }); } } required={true} />
                        <div className={styles.col_6}>
                  <TextField label="AX Number" required={this.state.approvepurchaseOptions=="Approve"?true:false}
                    value={this.state.axNumber}
                    onChanged={newVal => {
                      newVal && newVal.length > 0
                        ? this.setState({
                          axNumber: newVal,
                          erroraxNumber: ""
                        })
                        : this.state.approvepurchaseOptions=="Approve"?this.setState({
                          axNumber: newVal,
                          erroraxNumber:
                            "AX Number is required"
                        }):"";
                    }}
                    errorMessage={this.state.erroraxNumber}></TextField></div>
                        </div>
                       
                        
                        </>


                      )
                    
                     }
                  

                     })
                     :""
                    }
                 
                </div>
              </div>
            </div>
               <div className={classnames(styles.row, "reportBtn")}>
                 <div className={classnames(styles.col_3, styles.btnCancel)}>
                   <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
                 </div>
     
                     <div className={styles.col_3}>
                   <PrimaryButton className={styles.btnSubmit} text="Submit" id={"adminsubmit"} onClick={(e)=>this.state.pendingAtStage=="HSE Stage"?this.submitFormHSE.call(this,e):this.state.pendingAtStage=="Purchasing Stage"?this.submitFormPurchasing.call(this,e):this.state.pendingAtStage=="Finance Stage"?this.submitFormFinance.call(this,e):""} />
                 </div>
               </div>
             </div>:
          this.state.hideHseApprove && this.state.isHSEGroupApprover&&this.queryMode!="view" ?   
          <div >
           {
             StatusSummary&&StatusSummary.length>0? <div>
             <div className={styles.row}>
         <h3>History Contractor</h3>
           <div className={styles.col_12}>
             <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead ><tbody>
             {
               StatusSummary.map((rowDet)=>{
                 if(rowDet)
                 {
                  rowDet=rowDet.split('~');
             return(<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>)
               
               }})
             }
             
             </tbody></table>
           </div>
         </div><hr/></div>:""
           }
          <div className={classnames(styles.row, "radiostatus")}>
            <div className={styles.col_6}>
              <Label className={styles.divalign} required>Select status</Label>
              <ChoiceGroup defaultSelectedKey="A" options={this.options} onChange={(e, option) => {
option.text=="Approve"?
                this.setState({ approveOptions: option.text,hidereturnBox:true }):  this.setState({ approveOptions: option.text,hidereturnBox:false })
              }} required={true} />
            </div>
            <div className={styles.col_6} hidden={this.state.hidereturnBox}>
            <TextField label="Comments" required={true}
                value={this.state.returncomments}
                onChanged={newVal => {
                  newVal && newVal.length > 0
                    ? this.setState({
                      returncomments: newVal,
                      errorreturncomments: ""
                    })
                    : this.setState({
                      returncomments: newVal,
                      errorreturncomments: "Comments is required"
                    })
                }}
                errorMessage={this.state.errorreturncomments}></TextField>
            </div>
            </div>
          <div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div>

                <div className={styles.col_3}>
              <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitFormHSE} />
            </div>
          </div>
        </div>:this.state.isHSEGroupApprover&&this.queryMode=="view"?          <><div className={styles.row}>
                            <h3>History Contractor</h3>
                            <div className={styles.col_12}>
                            <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead><tbody>
                              {StatusSummary.map((rowDet) => {
                                if (rowDet) {
                                  rowDet = rowDet.split('~');
                                  return (<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>);

                                }
                              })}

                            </tbody></table>
                          </div>
                        </div><hr/> <div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)} style={{textAlign:"center"}}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div></div></>:""}

          {/* {Purchasing State Level        } */}
          {
          this.state.hidePurchasingApprove && this.state.isPurchasingGroupApprover&&this.queryMode!="view" ?   <div>
          <div className={styles.row}>
          <h3>History Contractor</h3>
            <div className={styles.col_12}>
              <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead ><tbody>
              {
                StatusSummary.map((rowDet)=>{
                  if(rowDet)
                  {
                   rowDet=rowDet.split('~');
              return(<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>)
                
                }})
              }
              
              </tbody></table>
            </div>
          </div><hr/>
          <div className={classnames(styles.row, "radiostatus")}>
            <div className={styles.col_6}>
              <Label className={styles.divalign} required>Select status</Label>
              <ChoiceGroup defaultSelectedKey="A" options={this.purchaseoptions} onChange={(e, option) => {

                this.setState({ approvepurchaseOptions: option.text })
              }} required={true} />
            </div></div>
          <div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div>

            {this.state.PurchasingStatus=="Pending"||this.state.PurchasingStatus=="In-Progress"?  <div className={styles.col_3}>
              <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitFormPurchasing} />
            </div>:""}
          </div>
        </div>:this.state.isPurchasingGroupApprover&&this.queryMode=="view"?          <><div className={styles.row}>
        <h3>History Contractor</h3>
                          <div className={styles.col_12}>
                            <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead><tbody>
                              {StatusSummary.map((rowDet) => {
                                if (rowDet) {
                                  rowDet = rowDet.split('~');
                                  return (<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>);

                                }
                              })}

                            </tbody></table>
                          </div>
                        </div> <hr/><div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)} style={{textAlign:"center"}}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div></div></>:""
          }

          {/* {Finance State Level        } */}
          {
          this.state.hideFinanceApprove && this.state.isFinanceGroupApprover&&this.queryMode!="view" ?  <div>
          <div className={styles.row}>
          <h3>History Contractor</h3>
              <div className={styles.col_12}>
              <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead ><tbody>
              {
                StatusSummary.map((rowDet)=>{
                  if(rowDet)
                  {
                   rowDet=rowDet.split('~');
              return(<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>)
                
                }})
              }
              </tbody></table>
            </div>
          </div><hr/>
          <div className={classnames(styles.row, "radiostatus")}>
            <div className={styles.col_6}>
              <Label className={styles.divalign} required>Select status</Label>
              <ChoiceGroup defaultSelectedKey="A" options={this.purchaseoptions} onChange={(e, option) => {

                this.setState({ approvepurchaseOptions: option.text })
              }} required={true} />
            </div>
            <div className={styles.col_6}>
              <TextField label="AX Number" required={this.state.approvepurchaseOptions=="Approve"?true:false}
                value={this.state.axNumber}
                onChanged={newVal => {
                  newVal && newVal.length > 0
                    ? this.setState({
                      axNumber: newVal,
                      erroraxNumber: ""
                    })
                    : this.state.approvepurchaseOptions=="Approve"?this.setState({
                      axNumber: newVal,
                      erroraxNumber:
                        "AX Number is required"
                    }):"";
                }}
                errorMessage={this.state.erroraxNumber}></TextField></div>
          </div>
          <div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div>

            {this.state.FinanceStatus=="Pending"||this.state.FinanceStatus=="In-Progress"?   <div className={styles.col_3}>
              <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitFormFinance} />
            </div>:""}
          </div>
        </div>:this.state.isFinanceGroupApprover&&this.queryMode=="view"?          <><div className={styles.row}>
        <h3>History Contractor</h3><div className={styles.col_12}>
                            <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead><tbody>
                              {StatusSummary.map((rowDet) => {
                                if (rowDet) {
                                  rowDet = rowDet.split('~');
                                  return (<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>);

                                }
                              })}

                            </tbody></table>
                          </div>
                        </div><hr/> <div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)} style={{textAlign:"center"}}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div></div></>:""
          }

          </div>:this.state.isnotApprover&&this.queryMode=="view"?   <>{StatusSummary.length>0?<div className={styles.row}>
          <h3>History Contractor</h3> <div className={styles.col_12}>
                            <table className={styles.table}><thead><tr><th>Status</th><th>Level</th><th>Date</th></tr></thead><tbody>
                              {StatusSummary.map((rowDet) => {
                                if (rowDet) {
                                  rowDet = rowDet.split('~');
                                  return (<tr><td>{rowDet[2]}</td><td>{rowDet[0]}</td><td>{rowDet[3]}</td></tr>);

                                }
                              })}

                            </tbody></table>
                          </div>
                        </div> :""} <hr/><div className={classnames(styles.row, "reportBtn")}>
            <div className={classnames(styles.col_3, styles.btnCancel)} style={{textAlign:"center"}}>
              <PrimaryButton className={styles.btnDraft} text="Cancel" onClick={this.cancelForm} />
            </div></div></>:""
  }
         
        </div>
      
      </div>
    );
  }
}
