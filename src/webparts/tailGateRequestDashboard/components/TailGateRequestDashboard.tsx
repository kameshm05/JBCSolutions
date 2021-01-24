import * as React from 'react';
import styles from './TailGateRequestDashboard.module.scss';
import { ITailGateRequestDashboardProps } from './ITailGateRequestDashboardProps';
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
import { Label } from 'office-ui-fabric-react/lib/Label';
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
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";
import { fetch as fetchPolyfill } from "whatwg-fetch";
import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  ContextualMenu,
  Toggle,
  Modal,
  IDragOptions,
  IIconProps,
} from 'office-ui-fabric-react';
import { IconButton, PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { AgGridColumn, AgGridReact } from 'ag-grid-react';

import 'ag-grid-community/dist/styles/ag-grid.css';
import 'ag-grid-community/dist/styles/ag-theme-alpine.css';
import { IStackProps, Link, Stack } from 'office-ui-fabric-react';
import {ITailGateRequestDashboardState} from './ITailGateRequestDashboardState'
import classnames from 'classnames';
import '../../../ExternalRef/style.css';
var imageSource='../../../ExternalRef/Images/icon-admin-edit.png'
export interface EditFormState {
  isTaskView: boolean,
  isEditView: boolean
}
var  _allItems: any[] = [];
var DraftDetails: any[]=[];
var completeDetails: any[]=[]; 
var readDetails:any[]=[];
var CurrentUserID;
export default class TailGateRequestDashboard extends React.Component<ITailGateRequestDashboardProps,ITailGateRequestDashboardState> {

 

  public state;
  listName: any = "TailgateTasksActivity";
  contractListName="ContractorsManagement";
  private _columns: IColumn[];
  private _Readcolumns:IColumn[];
  private _Activecolumns:IColumn[];
  private ActivecolumnDefs:any;
  private CompletecolumnDefs:any;
  private ReadcolumnDefs:any;
  private DraftcolumnDefs:any;
  options: IChoiceGroupOption[]
  DeleteIcon: IIconProps = { iconName: 'Delete' };
  constructor(props:ITailGateRequestDashboardProps) {
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
    this.CompletecolumnDefs= [
      {
        headerName: 'Task Identifier',
        field: 'TaskIdentifier',       
        sortable:true,filter:true,
        cellStyle: {fontWeight:'bold',cursor: 'pointer',color:'#c70808eb' },
        onCellClicked:function(e){
         window.location.href=e.data.redirectUrl
        }
     
      },
      { headerName: "Process Type", field: "ProcessType" ,sortable:true,filter:true}, 
      { headerName: "Overall Status", field: "OverallStatus",sortable:true,filter:true },   
      { headerName: "Requester", field: "Requester",sortable:true,filter:true },  
      { headerName: "Request Date", field: "RequestedDate",sortable:true,filter:true },
      { headerName: "RedirectUrl", field: "RedirectUrl",sortable:true,filter:true ,hide:true}
    ]

    this.ActivecolumnDefs= [
      {
        headerName: 'Task Identifier',
        field: 'TaskIdentifier',
        cellStyle: {fontWeight:'bold',cursor: 'pointer',color:'#c70808eb' },
        onCellClicked:function(e){
         window.location.href=e.data.redirectUrl
        },
        sortable:true,filter:true
     
      },
      { headerName: "Process Type", field: "ProcessType" ,sortable:true,filter:true},  
      { headerName: "Requester", field: "Requester",sortable:true,filter:true }, 
      { headerName: "Assigned Date", field: "assignedDate",sortable:true,filter:true },   
      { headerName: "Request Date", field: "RequestedDate",sortable:true,filter:true },
      { headerName: "RedirectUrl", field: "RedirectUrl",sortable:true,filter:true ,hide:true}
    ]

    this.ReadcolumnDefs= [
      {
        headerName: 'Task Identifier',
        field: 'TaskIdentifier',
        cellRenderer:function(params){
          var urlLen=params.data.redirectUrl.split('~')
          if(urlLen.length>1)
          {
            var resultElement = document.createElement("span");
            var aElement = document.createElement("a");
            var imgAElemet=document.createElement("a");
                var imageElement = document.createElement("span"); 
                imageElement.classList.add('tableSignOff');
                aElement.href=urlLen[0]; 
                imgAElemet.href=urlLen[1]; 
                imgAElemet.appendChild(imageElement);               
                aElement.innerText=params.data.TaskIdentifier;
                resultElement.appendChild(aElement);
                resultElement.appendChild(imgAElemet); 
             
            return resultElement;
          }
          else
          {  
            var resultElement = document.createElement("span");
          var aElement = document.createElement("a");
          aElement.href=urlLen[0]; 
          aElement.innerText=params.data.TaskIdentifier;
          resultElement.appendChild(aElement);
          return resultElement; 
          }

        },
        cellStyle: {fontWeight:'bold',cursor: 'pointer',color:'#c70808eb' },
        sortable:true,filter:true
      },
      { headerName: "Process Type", field: "ProcessType" ,sortable:true,filter:true},  
      { headerName: "Requester", field: "Requester",sortable:true,filter:true }, 
      { headerName: "Pending At", field: "PendingAt",sortable:true,filter:true },   
      { headerName: "Request Date", field: "RequestedDate",sortable:true,filter:true },
      { headerName: "RedirectUrl", field: "RedirectUrl",sortable:true,filter:true ,hide:true}

    ]
    this.DraftcolumnDefs= [
      {
        headerName: 'Task Identifier',
        field: 'TaskIdentifier',
       
        cellStyle: {fontWeight:'bold',cursor: 'pointer',color:'#c70808eb' },
        onCellClicked:function(e){
         window.location.href=e.data.redirectUrl
        },
        sortable:true,filter:true
     
      },
      { headerName: "Process Type", field: "ProcessType" ,sortable:true,filter:true},  
      { headerName: "Requester", field: "Requester",sortable:true,filter:true },  
      { headerName: "Request Date", field: "RequestedDate",sortable:true,filter:true },
      { headerName: "RedirectUrl", field: "RedirectUrl",sortable:true,filter:true ,hide:true}

    ]
   
    this.state = {
      getActiveDataDetails: [],
      get_draftDetails: [],
      get_completeDetails: [],
      get_readonlyDetails: [],
      get_Active_Paged_array:[],
      get_Draft_Paged_array:[],
      get_Completed_Paged_array:[],
      get_Read_Paged_array:[],
      filterTaskDetails: "",
      filter_draftDetails: "",
      filter_completeDetails: "",
      filter_readonlyDetails: "",
      isTaskView: true,
      isEditView: false,
      Topic: "",
      description: "",
      fileDetails: [],
      ApprovalModal:true,
      SignOffModal:true,
      ItemId:0,
      approveStatus:"Approve",
      chksignOffStatus:false,
      comments:"",
      errorcomments:"",
      StatusSummary:"",
      fetchApprovers:[],
      fetchSignOffUsers:[],
      btnsReadonly:false,
      EditModel:true,
      errortopicValue: "",
      errordescriptionValue: "",
      filePickerResult:[],
      getUsers:[],
      allpeoplePicker_User:[],
      getApprovepeoplePicker_User:[],
      getSignOffUser:[],
      allpeoplePicker2_User:[],
      errorSignoffUsers:false,
      errorapproverUsers:false,
      topicValue:"",
      descriptionValue:"",
      errorfileAttach:"",
      newFiles:[],
      removedFiles:[],
      isAdmin:false
      
    }
  


   
    this.init();

  }

  public init=()=>{
    _allItems=[];
    this.getGroupDetails();
   


  }
public getGroupDetails=()=>{
  let groups = sp.web.currentUser.groups().then((grpDetails)=>{
    grpDetails.map((eachGroup,idx)=>{
      if(eachGroup.Title=="AdminTeam")
      {
        this.setState({isAdmin:true});
        this.state.isAdmin=true;
      }
      if(grpDetails.length<=idx+1)
        this.getCurrentUserDetails();
    });
    if(grpDetails.length==0)
    this.getCurrentUserDetails();
  })
}
public getCurrentUserDetails=()=>{
  sp.web.currentUser.get().then((UserId) => {
    CurrentUserID=UserId['Id'];
    this.getallDatas(UserId);
    this.getallDraftDetails(UserId);
    this.getallCompleteDetails(UserId);
    this.getallReadOnlyDetails(UserId);
  });
}
public getCurrentUserReadonlyfilter=(UserId)=>{

  var groupFilter="AuthorId eq '" + UserId['Id'] + "' and FinanceStatus ne 'Approved' and Status ne 'Draft' and HSEStatus ne 'Returned' or HSEStatus eq 'Pending' or PurchasingStatus eq 'Pending' or FinanceStatus eq 'Pending'"

  this.getContractReadOnlyDetails(groupFilter,UserId);

    // let groups =  sp.web.currentUser.groups().then((grpDetails)=>{
    //   var groupFilter="";
    //   grpDetails.map((eachGroup,idx)=>{
    //     if(eachGroup.Title=="HSE_Approver")
    //     {
    //       if(idx==0)
    //       groupFilter="HSEStatus eq 'Approved' and Status ne 'Draft' and HSEStatus ne 'Returned' or ";
    //       else
    //       groupFilter=groupFilter+"HSEStatus eq 'Approved' and Status ne 'Draft' and HSEStatus ne 'Returned' or ";
          
    //     }
    //  else if(eachGroup.Title=="Purchasing")
    //     {
    //       if(idx==0)
    //       groupFilter="HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus eq 'Pending' and Status ne 'Draft' and HSEStatus ne 'Returned' or ";
    //       else
    //       groupFilter=groupFilter+"HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus eq 'Pending' and Status ne 'Draft' and HSEStatus ne 'Returned' or ";
    //     }
    //     else if(eachGroup.Title=="Finance")
    //     {
    //       if(idx==0)
    //       groupFilter="HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus ne 'Pending' and Status ne 'Draft' and HSEStatus ne 'Returned' or ";
    //       else
    //       groupFilter=groupFilter+"HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus ne 'Pending' and Status ne 'Draft' and HSEStatus ne 'Returned' or ";
    //     }    
    //   });

    //   if(groupFilter!="")
    //   {
    //     var groupFilter=groupFilter+"AuthorId eq '" + UserId['Id'] + "' and FinanceStatus ne 'Pending' and Status ne 'Draft' and HSEStatus ne 'Returned' or "
    //     var orindex=groupFilter.lastIndexOf('or');
    //     groupFilter=groupFilter.substring(0,orindex-1)
    //     this.getContractReadOnlyDetails(groupFilter);

    //   }
    //   else
    //   {
    //     var groupFilter="AuthorId eq '" + UserId['Id'] + "' and FlowStatus ne 'Finance Approved' and Status ne 'Draft' and HSEStatus ne 'Returned' or "
    //     var orindex=groupFilter.lastIndexOf('or');
    //     groupFilter=groupFilter.substring(0,orindex-1)
    //     this.getContractReadOnlyDetails(groupFilter);
    //   }

    // });
  

}
public getCurrentUserCompletefilter=(UserId)=>{
  var groupFilter="";
  // var groupFilter="HSEStatus eq 'Approved' and PurchasingStatus eq 'Approved' and FinanceStatus eq 'Approved' and Status ne 'Draft' or (AuthorId eq '" + UserId['Id'] + "' and HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus ne 'Pending' and Status ne 'Draft' or (AuthorId eq '" + UserId['Id'] + "' and  HSEStatus eq 'Returned'))";
  // this.getContractcompleteDetails(groupFilter);

  let groups =  sp.web.currentUser.groups().then((grpDetails)=>{
    var groupFilter="";
  var finalGroups=  grpDetails.filter((eachGroup,idx)=>{
    return eachGroup.Title=="HSE_Approver"||eachGroup.Title=="Purchasing"||eachGroup.Title=="Finance"||eachGroup.Title=="AdminTeam" 
    });
    if(finalGroups.length>0)
    {
     
        groupFilter="HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus ne 'Pending' and Status ne 'Draft' or  HSEStatus eq 'Returned'";  
    }

    if(groupFilter!="")
    {
      this.getContractcompleteDetails(groupFilter);
    }   
    else 
    {
       groupFilter=groupFilter+"AuthorId eq '" + UserId['Id'] + "' and HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and FinanceStatus ne 'Pending' and Status ne 'Draft' or (AuthorId eq '" + UserId['Id'] + "' and  HSEStatus eq 'Returned')"
      this.getContractcompleteDetails(groupFilter);
    }
  });
}

  public getContractActiveDatas=(UserId)=>{
    
    var groupFilter="";
   
 
      let groups =  sp.web.currentUser.groups().then((grpDetails)=>{
      
        grpDetails.map((eachGroup,idx)=>{
          if(eachGroup.Title=="HSE_Approver")
          {
            if(idx==0)
            groupFilter="Status eq 'Submit' and HSEStatus eq 'Pending' or ";
            else
            groupFilter=groupFilter+"Status eq 'Submit' and HSEStatus eq 'Pending' or ";
            
          }
        if(eachGroup.Title=="Purchasing")
          {
            if(idx==0)
            groupFilter="Status eq 'Submit' and HSEStatus eq 'Approved' and (PurchasingStatus eq 'Pending) or ";
            else
            groupFilter=groupFilter+"Status eq 'Submit' and HSEStatus eq 'Approved' and (PurchasingStatus eq 'Pending') or ";
          }
          if(eachGroup.Title=="Finance")
          {
            if(idx==0)
            groupFilter="Status eq 'Submit' and HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and  (FinanceStatus eq 'Pending') or ";
            else
            groupFilter=groupFilter+"Status eq 'Submit' and HSEStatus eq 'Approved' and PurchasingStatus ne 'Pending' and (FinanceStatus eq 'Pending') or ";
          }
          if(eachGroup.Title=="AdminTeam") 
          {
            if(idx==0)
             groupFilter="HSEStatus eq 'Pending' or PurchasingStatus eq 'Pending' or FinanceStatus eq 'Pending' and Status ne 'Draft' and  HSEStatus ne 'Returned' or ";
            else
            groupFilter=groupFilter+"HSEStatus eq 'Pending' or PurchasingStatus eq 'Pending' or FinanceStatus eq 'Pending' and Status ne 'Draft' and  HSEStatus ne 'Returned' or ";
          }    
        });
  
        if(groupFilter!="")
        {
          var orindex=groupFilter.lastIndexOf('or');
          groupFilter=groupFilter.substring(0,orindex-1)
          this.getContractActiveDetails(groupFilter);
        }
        else
        {
           groupFilter="HSEStatus eq 'Pending' or PurchasingStatus eq 'Pending' or FinanceStatus eq 'Pending' and Status ne 'Draft' and  HSEStatus ne 'Returned'";
          this.getContractActiveDetails(groupFilter);
        }
        // else
        // {
        //   var groupFilter="HSEStatus eq 'Pending' or PurchasingStatus eq 'Pending' or FinanceStatus eq 'Pending' and Status ne 'Draft'   and  HSEStatus ne 'Returned'"
        //   this.setState({ getActiveDataDetails: _allItems,get_Active_Paged_array:_allItems});
        // }
      });

  }

  public getContractActiveDetails=(groupFilter)=>{
    sp.web.lists.getByTitle(this.contractListName).items.select("*,Author/Title").expand("Author").filter(groupFilter).orderBy("Created",false).getAll().then((HSEItem: any) => {
      let modeObj: any;
     for(let i=0;i<HSEItem.length;i++)
     {
      if(HSEItem[i].Level)
      {
        var currentUserName=this.props.spcontext.pageContext.user.displayName;
        var HSEPType=HSEItem[i].HSEProcessType;
        var PurchasingPType=HSEItem[i].PurchasingProcessType;
        var FinancePType=HSEItem[i].FinanceProcessType;
      
        var responseHSEIdx=0;
        var responsePurAppIdx=0; 
        var responsePurInIdx=0;
        var responseFinAppIdx=0;
        var responseFinInIdx=0;

        if(HSEItem[i].HSEStatus=="Pending")
        {
          if(HSEPType.indexOf('HSE~User~'+currentUserName)>=0)
          {
            var chkCon="HSE~"+currentUserName+"~Approved";
             responseHSEIdx=HSEItem[i].Level.indexOf(chkCon);
          }
          else if(HSEPType.indexOf('HSE~Group')>=0)
          {
            var chkCon="HSE~"+currentUserName+"~Approved";
             responseHSEIdx=HSEItem[i].Level.indexOf(chkCon);
          }
        }
        else if(HSEItem[i].PurchasingStatus=="Pending")
        {
          if(PurchasingPType.indexOf('Purchasing~User~'+currentUserName)>=0)
          {
            var chkConApprove="Purchasing~"+currentUserName+"~Approved";
            var chkConIn="Purchasing~"+currentUserName+"~In-Progress";

             responsePurAppIdx=HSEItem[i].Level.indexOf(chkConApprove);
             responsePurInIdx=HSEItem[i].Level.indexOf(chkConIn);

          }
          else if(PurchasingPType.indexOf('Purchasing~Group')>=0)
          {
            var chkConApprove="Purchasing~"+currentUserName+"~Approved";
            var chkConIn="Purchasing~"+currentUserName+"~In-Progress";

             responsePurAppIdx=HSEItem[i].Level.indexOf(chkConApprove);
             responsePurInIdx=HSEItem[i].Level.indexOf(chkConIn);
          }
        }
        else if(HSEItem[i].FinanceStatus=="Pending")
        {
          if(FinancePType.indexOf('Finance~User~'+currentUserName)>=0)
          {      
            var chkConApprove="Finance~"+currentUserName+"~Approved";
            var chkConIn="Finance~"+currentUserName+"~In-Progress";

             responseFinAppIdx=HSEItem[i].Level.indexOf(chkConApprove);
             responseFinInIdx=HSEItem[i].Level.indexOf(chkConIn);
          }
          else if(FinancePType.indexOf('Finance~Group')>=0)
          {
             chkConApprove="Finance~"+currentUserName+"~Approved";
             chkConIn="Finance~"+currentUserName+"~In-Progress";

             responseFinAppIdx=HSEItem[i].Level.indexOf(chkConApprove);
             responseFinInIdx=HSEItem[i].Level.indexOf(chkConIn);
          }
        }

        
      }
      if(HSEItem[i].FinanceStatus!="Approved"&& this.state.isAdmin)
      {
        modeObj=this.props.siteUrl+"/SitePages/ApprovePage.aspx?IDCO="+HSEItem[i].ID.toString()+"&CMode=editAdmin"
        var arritems = {
          TaskIdentifier: HSEItem[i].ContractorsName,
          ProcessType: HSEItem[i]["Title"],
          Requester:  HSEItem[i]["Author"]["Title"],
          assignedDate:HSEItem[i].AssignedDate?HSEItem[i].AssignedDate:"",
          RequestedDate:new Date(HSEItem[i].Created).toLocaleDateString(),
          redirectUrl:modeObj
        };
        _allItems = _allItems.concat(arritems);
      }
      else if(responseHSEIdx==-1||(responsePurAppIdx==-1&&responsePurInIdx==-1)||(responseFinAppIdx==-1&&responseFinInIdx==-1))
      {
        modeObj=this.props.siteUrl+"/SitePages/ApprovePage.aspx?IDCO="+HSEItem[i].ID.toString()+""
        var arritems = {
          TaskIdentifier: HSEItem[i].ContractorsName,
          ProcessType: HSEItem[i]["Title"],
          Requester:  HSEItem[i]["Author"]["Title"],
          assignedDate:HSEItem[i].AssignedDate?HSEItem[i].AssignedDate:"",
          RequestedDate:new Date(HSEItem[i].Created).toLocaleDateString(),
          redirectUrl:modeObj
        };
        _allItems = _allItems.concat(arritems);
      }
     }
     _allItems=this.sortByDate(_allItems);


   
     this.setState({ getActiveDataDetails: _allItems,get_Active_Paged_array:_allItems });
    });
  }
  
  sortByDate(arr) {
    arr.sort(function(a,b){
      return Number(new Date(b.RequestedDate)) - Number(new Date(a.RequestedDate));
    });

    return arr;
  }

  public getContractDraftDetails=(UserId)=>{
    sp.web.lists.getByTitle(this.contractListName).items.select("*,Author/Title").expand("Author").filter("AuthorId eq '" + UserId['Id'] + "' and Status eq 'Draft'").orderBy("Created",false).getAll().then((DraftItem: any) => {
      let modeObj: any;
     for(let i=0;i<DraftItem.length;i++)
     {

      modeObj = this.props.siteUrl+"/SitePages/ContractManagementRequest.aspx?IDCO="+DraftItem[i].ID.toString()+""
      var arritems = {
        TaskIdentifier: DraftItem[i].ContractorsName,
        ProcessType: DraftItem[i]["Title"],
      
        Requester:  DraftItem[i]["Author"]["Title"],
        RequestedDate:  new Date(DraftItem[i].Created).toLocaleDateString(),
        redirectUrl:modeObj
      };
       DraftDetails.push(arritems);
     }

     DraftDetails=this.sortByDate(DraftDetails);



     this.setState({ get_draftDetails: DraftDetails,get_Draft_Paged_array:DraftDetails });
    
    });
  }
  public getContractcompleteDetails=(filter)=>{
    sp.web.lists.getByTitle(this.contractListName).items.select("*,Author/Title,Author/Id").expand("Author").filter(filter).orderBy("Created",false).getAll().then((ContractItem: any) => {
      let modeObj: any;
     for(let i=0;i<ContractItem.length;i++)
     {
     
       if(ContractItem[i].HSEStatus=="Returned"&&ContractItem[i].Author.Id==CurrentUserID)
       {
         modeObj=this.props.siteUrl+"/SitePages/ContractManagementRequest.aspx?IDCO="+ContractItem[i].ID.toString()+""

       }
       else
       {
         modeObj=this.props.siteUrl+"/SitePages/ViewPage.aspx?IDCO="+ContractItem[i].ID.toString()+"&CMode=View"
       }


      var arritems = {
        TaskIdentifier: ContractItem[i].ContractorsName,
        ProcessType: ContractItem[i]["Title"],
      
        Requester:  ContractItem[i]["Author"]["Title"],
        OverallStatus:ContractItem[i].HSEStatus=="Returned"?"Returned":ContractItem[i].FinanceStatus != 'Pending'&&ContractItem[i].FinanceStatus?ContractItem[i].FinanceStatus:"-",
        RequestedDate:  new Date(ContractItem[i].Created).toLocaleDateString(),
        redirectUrl:modeObj
      };
      completeDetails.push(arritems);
     }
     completeDetails=this.sortByDate(completeDetails);
     this.setState({ get_completeDetails: completeDetails,get_Completed_Paged_array:completeDetails });
    });
  }
  public getContractReadOnlyDetails=(filter,UserId)=>{

   
    sp.web.lists.getByTitle(this.contractListName).items.select("*,Author/Title").expand("Author").filter(filter).orderBy("Created",false).getAll().then((ContractItem: any) => {
      let modeObj: any;
     for(let i=0;i<ContractItem.length;i++)
     {
       if(ContractItem[i].Level)
       {

        var currentUserName=this.props.spcontext.pageContext.user.displayName;
        var HSEPType=ContractItem[i].HSEProcessType;
        var PurchasingPType=ContractItem[i].PurchasingProcessType;
        var FinancePType=ContractItem[i].FinanceProcessType;


        var responseHSEIdx=-1;
        var responsePurAppIdx=-1; 
        var responsePurInIdx=-1;
        var responseFinAppIdx=-1;
        var responseFinInIdx=-1;
        
          if(HSEPType.indexOf('HSE~User~'+currentUserName)>=0)
          {
            var chkCon="HSE~"+currentUserName+"~Approved";
              responseHSEIdx=ContractItem[i].Level.indexOf(chkCon);
          }
          else if(HSEPType.indexOf('HSE~Group')>=0)
          {
            var chkCon="HSE~"+currentUserName+"~Approved";
              responseHSEIdx=ContractItem[i].Level.indexOf(chkCon);
          }
       
        
          if(PurchasingPType.indexOf('Purchasing~User~'+currentUserName)>=0)
          {
            var chkConApprove="Purchasing~"+currentUserName+"~Approved";
            var chkConIn="Purchasing~"+currentUserName+"~In-Progress";
  
              responsePurAppIdx=ContractItem[i].Level.indexOf(chkConApprove);
              responsePurInIdx=ContractItem[i].Level.indexOf(chkConIn);
  
          }
          else if(PurchasingPType.indexOf('Purchasing~Group')>=0)
          {
            var chkConApprove="Purchasing~"+currentUserName+"~Approved";
            var chkConIn="Purchasing~"+currentUserName+"~In-Progress";
  
              responsePurAppIdx=ContractItem[i].Level.indexOf(chkConApprove);
              responsePurInIdx=ContractItem[i].Level.indexOf(chkConIn);
          }
        
          if(FinancePType.indexOf('Finance~User~'+currentUserName)>=0)
          {      
            var chkConApprove="Finance~"+currentUserName+"~Approved";
            var chkConIn="Finance~"+currentUserName+"~In-Progress";
  
              responseFinAppIdx=ContractItem[i].Level.indexOf(chkConApprove);
              responseFinInIdx=ContractItem[i].Level.indexOf(chkConIn);
          }
          else if(FinancePType.indexOf('Finance~Group')>=0)
          {
            var chkConApprove="Finance~"+currentUserName+"~Approved";
            var chkConIn="Finance~"+currentUserName+"~In-Progress";
  
              responseFinAppIdx=ContractItem[i].Level.indexOf(chkConApprove);
              responseFinInIdx=ContractItem[i].Level.indexOf(chkConIn);
          }
        
  
        
          if(responseHSEIdx>=0||responsePurAppIdx>=0||responsePurInIdx>=0||responseFinAppIdx>=0||responseFinInIdx>=0||(ContractItem[i].Author.Title==UserId.Title&&ContractItem[i].FinanceStatus!="Approved")) 
          {
            modeObj=this.props.siteUrl+"/SitePages/ViewPage.aspx?IDCO="+ContractItem[i].ID.toString()+"&CMode=View"
            var arritems = {
              TaskIdentifier: ContractItem[i].ContractorsName,
              ProcessType: ContractItem[i]["Title"],
              
              Requester:  ContractItem[i]["Author"]["Title"],
              PendingAt:ContractItem[i].PendingAt?ContractItem[i].PendingAt:"",
              RequestedDate:  new Date(ContractItem[i].Created).toLocaleDateString(),
              redirectUrl:modeObj
              
            };
            readDetails.push(arritems);
          }  
          
  

       }
      
     }

     readDetails=this.sortByDate(readDetails);


     this.setState({ get_readonlyDetails: readDetails,get_Read_Paged_array:readDetails});
    });
  }

  public getallDatas = (UserId) =>{
    if(!this.state.isAdmin)
    {
      var query=sp.web.lists.getByTitle(this.listName).items.select("AssignedDate,ProcessType,TaskIdentifier,Status,Approvers/Id,Approvers/Title,Signoffs/Name,Signoffs/Title,Author/Title,Signoffs/Id,Created,ID,ApprovalSummary,SignOffStatus").expand("Approvers,Signoffs,Author").filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Submit' or (SignoffsId eq '" + UserId['Id'] + "' and SignOffStatus eq 'SignOff Pending')").orderBy("Created",false) 
    }
    else
    {
      var query=sp.web.lists.getByTitle(this.listName).items.select("AssignedDate,ProcessType,TaskIdentifier,Status,Approvers/Id,Approvers/Title,Signoffs/Name,Signoffs/Title,Author/Title,Signoffs/Id,Created,ID,ApprovalSummary,SignOffStatus").expand("Approvers,Signoffs,Author").filter("SignOffStatus ne 'SignOff Completed' and Status ne 'Returned' and Status ne 'Draft'").orderBy("Created",false)
    }
    query.get().then((Items: any) => {
          let modeObj: any;
          _allItems=[];
          for (var i = 0; i < Items.length; i++) {
           if(Items[i].ApprovalSummary)
           {
             var userDisplayName=this.props.spcontext.pageContext.user.displayName;
             var checkCondition=userDisplayName+"~Approved";
            var responseIdx=Items[i].ApprovalSummary.indexOf(checkCondition);
            if(responseIdx==-1 && Items[i].SignOffStatus!="SignOff Pending"||(this.state.isAdmin&& Items[i].SignOffStatus!="SignOff Pending"))
            {
              if(this.state.isAdmin)
              modeObj =  this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+Items[i].ID.toString()+"&CMode=EditAdmin"
              else
              modeObj =  this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+Items[i].ID.toString()+"&CMode=Edit"
              var arritems = {
                TaskIdentifier: Items[i].TaskIdentifier,
                ProcessType: Items[i]["ProcessType"],     
                Requester:  Items[i]["Author"]["Title"],
                assignedDate:Items[i].AssignedDate?Items[i].AssignedDate:"",
                RequestedDate:  new Date(Items[i].Created).toLocaleDateString(),
                redirectUrl:modeObj
              };
              _allItems = _allItems.concat(arritems);
            }
            else if(Items[i].SignOffStatus=="SignOff Pending"||(this.state.isAdmin&& Items[i].SignOffStatus=="SignOff Pending"))
            {
              var checkCondition=userDisplayName+"~SignOff Completed";
              var responseIdx=Items[i].ApprovalSummary.indexOf(checkCondition);
              if(responseIdx==-1&&Items[i].SignOffStatus=="SignOff Pending"||(this.state.isAdmin&& Items[i].SignOffStatus=="SignOff Pending"))
              {
                if(this.state.isAdmin)
                modeObj=this.props.siteUrl+"/SitePages/ReportSignOff.aspx?SID="+Items[i].ID.toString()+"&CMode=EditAdmin"
                else
                modeObj=this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+Items[i].ID.toString()+"&CMode=Edit"
                var arritems = {
                  TaskIdentifier: Items[i].TaskIdentifier,
                  ProcessType: Items[i]["ProcessType"],  
                  Requester:  Items[i]["Author"]["Title"],
                  assignedDate:Items[i].AssignedDate?Items[i].AssignedDate:"",
                  RequestedDate:  new Date(Items[i].Created).toLocaleDateString(),
                  redirectUrl:modeObj
                };
                _allItems = _allItems.concat(arritems);
              }
            }
           }

          }
        if(Items.length<=i+1)
        this.getContractActiveDatas(UserId);          
        });
  }


  private getallDraftDetails(UserId) {
      sp.web.lists.getByTitle(this.listName).items.select("*,Author/Title,Author/Id,Created").expand("Author")
        .filter("AuthorId eq '" + UserId['Id'] + "' and Status eq 'Draft'").orderBy("Created",false).get().then((Items: any) => {
           DraftDetails=[];
          let editObj:any;
          for (var i = 0; i < Items.length; i++) {
          editObj =this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+Items[i].ID.toString()
            var arritems = {
           
              TaskIdentifier:Items[i].TaskIdentifier,
              ProcessType: Items[i]["ProcessType"],
              Requester:  Items[i]["Author"]["Title"],
              RequestedDate: new Date(Items[i].Created).toLocaleDateString(),
              redirectUrl:editObj
            };
            DraftDetails.push(arritems);
          }
          if(Items.length<=i+1)
          this.getContractDraftDetails(UserId);
        });
  }

  private getallCompleteDetails(UserId) {
    if(!this.state.isAdmin)
    {
      var query= sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title,Author/Title,Author/Id,Signoffs/Id,Created,SignOffStatus,Status").expand("Approvers,Author,Signoffs")
      .filter("AuthorId eq '" + UserId['Id'] + "' and (SignOffStatus eq 'SignOff Completed' or Status eq 'Returned' ) or ApproversId eq '" + UserId['Id'] + "' and (SignOffStatus eq 'SignOff Completed'  or Status eq 'Returned') or  SignoffsId eq '" + UserId['Id'] + "' and SignOffStatus eq 'SignOff Completed'").orderBy("Created",false)
    }
    else
    {
      var query= sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title,Author/Title,Author/Id,Signoffs/Id,Created,SignOffStatus,Status").expand("Approvers,Author,Signoffs").orderBy("Created",false)
    }
    query.get().then((Items: any) => {
          let modeObj: any;
          completeDetails=[]; 
          for (var i = 0; i < Items.length; i++) {
            if(Items[i].ApprovalSummary)
           {
             var userDisplayName=this.props.spcontext.pageContext.user.displayName;
             var checkCondition=userDisplayName+"~Approved";
            var responseIdx=Items[i].ApprovalSummary.indexOf(checkCondition);
            var checkSignoff=userDisplayName+"~SignOff Completed";
            var signOffIdx=Items[i].ApprovalSummary.indexOf(checkSignoff);
            var checkUSer=userDisplayName+"~New Request";
            var userIdx=Items[i].ApprovalSummary.indexOf(checkUSer);
            var returnstatus=userDisplayName+"~Returned";
            var returnIdx=Items[i].ApprovalSummary.indexOf(returnstatus);

            if((Items[i].SignOffStatus=="SignOff Completed"||Items[i].Status=="Returned"))
            {
             if(Items[i].Status=="Returned"&&Items[i].Author.Id==CurrentUserID)
             modeObj=this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+Items[i].ID.toString()
             else
             modeObj=this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+Items[i].ID.toString()+"&CMode=View"

              var arritems = {
                TaskIdentifier:Items[i].TaskIdentifier ,
                ProcessType: Items[i]["ProcessType"],
                Requester:  Items[i]["Author"]["Title"],
                OverallStatus:Items[i].Status=="Returned"?"Returned":Items[i].SignOffStatus=="SignOff Completed"?"SignOff Completed":"-",
                RequestedDate:  new Date(Items[i].Created).toLocaleDateString(),
                redirectUrl:modeObj
              };
              completeDetails.push(arritems);
            }

           }
          }
          if(Items.length<=i+1)
          this.getCurrentUserCompletefilter(UserId)
        });
  }

  private getallReadOnlyDetails(UserId) {
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title,Author/Title,Author/Id,Signoffs/Id,ApproversText").expand("Approvers,Author,Signoffs")
        .filter("AuthorId eq '" + UserId['Id'] + "' and (Status eq 'Submit' or Status eq 'Approved' and  SignOffStatus ne 'SignOff Completed' and  Status ne 'Returned') or (ApproversId eq '" + UserId['Id'] + "' and ( Status ne 'Draft' and SignOffStatus ne 'SignOff Completed' and  Status ne 'Returned' )) or ( SignoffsId eq '" + UserId['Id'] + "' and (SignOffStatus eq 'SignOff Pending' and  Status ne 'Returned'))").orderBy("Created",false).get().then((AllItems: any[]) => {
          let modeObj: any;
           readDetails=[];

          for (var i = 0; i < AllItems.length; i++) {
            var responseIdx=0;
            var signresponseIdx=0
            if(AllItems[i].ApproversId)
            {
              var approveIdx=AllItems[i].ApproversId.includes(UserId['Id'])
            }
            if(AllItems[i].SignoffsId)
            {
              var signIdx=AllItems[i].SignoffsId.includes(UserId['Id'])
            }
            if(AllItems[i].ApprovalSummary && (approveIdx||signIdx))
            {
              var userDisplayName=this.props.spcontext.pageContext.user.displayName;
              var checkCondition=userDisplayName+"~Approved";
              responseIdx=AllItems[i].ApprovalSummary.indexOf(checkCondition);
              var checkSignoff=userDisplayName+"~SignOff Completed";
               signresponseIdx=AllItems[i].ApprovalSummary.indexOf(checkSignoff);
           
            }
            if((responseIdx>=0||signresponseIdx>=0)&&AllItems[i].SignOffStatus!="SignOff Completed")
            {
              if(AllItems[i]["Author"]["Id"]==UserId['Id']&&AllItems[i].ApprovalStatus=="Approved"&&AllItems[i].SignOffStatus=="SignOff Pending")
              { 
                  modeObj=this.props.siteUrl + "/SitePages/TailGateNewRequest.aspx?RID=" + AllItems[i].ID.toString() + "&CMode=View~"+this.props.siteUrl + "/SitePages/ReportSignOff.aspx?SID=" + AllItems[i].ID.toString() + ""
              }
              else{
                modeObj=this.props.siteUrl+"/SitePages/TailGateNewRequest.aspx?RID="+AllItems[i].ID.toString()+"&CMode=View"
              }
              var arritems = {
                TaskIdentifier: AllItems[i].TaskIdentifier,
                ProcessType: AllItems[i]["ProcessType"],
                Requester:  AllItems[i]["Author"]["Title"],
                PendingAt:AllItems[i].PendingAt?AllItems[i].PendingAt:"",
                RequestedDate:  new Date(AllItems[i].Created).toLocaleDateString(),
                redirectUrl:modeObj
              };
              readDetails.push(arritems);
            }

           

               

          }
          if(AllItems.length<=i+1)
          this.getCurrentUserReadonlyfilter(UserId);
        });
  }

 
  public _alertClicked = (): void => {
    this.setState({
      ApprovalModal: true,
      Topic:"",
      description:"",
      fileDetails:[]
    });
  }
  public SubmitForm = (): void => {
    if (this.state.approveStatus=="Approve") {
      var UserName=this.props.spcontext.pageContext.user.displayName;
      var comments=this.state.comments;
      var date =new Date().toLocaleDateString();
      var setSummary=UserName+"~Approved~"+comments+"~"+date+"|";
      var finalSummary=this.state.StatusSummary+setSummary;
      var finalIdx=finalSummary.split('|');
      var calIDCOx=finalIdx.length-2;
      var approversLength=this.state.fetchApprovers.length;
      if(calIDCOx==approversLength)
      {
        var statusUpdate="Approved";
        var signOffUpdate="SignOff Pending";
        var approverstatus="Approved"
      }
      else
      {
        var statusUpdate="Submit";
        var approverstatus="Pending"
      }



      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.state.ItemId).update({
        ApprovalSummary: this.state.StatusSummary+setSummary,
        Status:statusUpdate,
        SignOffStatus:signOffUpdate,
        ApprovalStatus:approverstatus
      }).then(s => {
        alert("Request updated successfully...!!!");
        this.init();
        this.setState({   ApprovalModal: true });
      });
    } else if( this.state.approveStatus=="Return"){
      if(!this.state.comments.length)
      {
        this.setState({ errorcomments: "Comments is Required" });
      
      }
      else
      {
        var UserName=this.props.spcontext.pageContext.user.displayName;
        var comments=this.state.comments;
        var date =new Date().toLocaleDateString();
        var setSummary=UserName+"~Returned~"+comments+"~"+date+"|";
        var statusUpdate="Returned"
        sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.state.ItemId).update({
          ApprovalSummary: this.state.StatusSummary+setSummary,
          Status:statusUpdate,
          ApprovalStatus:"Returned"
        }).then(s => {
          alert("Request updated successfully...!!!");
          this.init();
          this.setState({   ApprovalModal: true });
        });
      }

    }

  }

  public SignoffForm=()=>{
    if (this.state.chksignOffStatus==true) {
      var UserName=this.props.spcontext.pageContext.user.displayName;
      var comments="";
      var date =new Date().toLocaleDateString();
      var setSummary=UserName+"~SignOff Completed~"+comments+"~"+date+"|";
      var finalSummary=this.state.StatusSummary+setSummary;
      var count = finalSummary.match(/SignOff Completed/g);
      var clean_count = !count ? false : count.length;
      var calIDCOx=count.length;
      var SignOffUsersLength=this.state.fetchSignOffUsers.length;
      if(calIDCOx==SignOffUsersLength)
      {
        var signOffUpdate="SignOff Completed"
      }
      else
      {
        var signOffUpdate="SignOff Pending"
      }



      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.state.ItemId).update({
        ApprovalSummary: this.state.StatusSummary+setSummary,
        SignOffStatus:signOffUpdate
      }).then(s => {
        alert("Request updated successfully...!!!");
        this.init();
        this.setState({   SignOffModal: true });
      });
    } else if( this.state.approveStatus=="Return"){
     

    }
  }
  public fileUploadCallback = (e) => {

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
      this.setState({ newFiles: [] });
      e.target.value = null;
    }
  }
  public SignOffpeoplechange = (event) => {
   
    if(event["length"]>0)
    {
      var resultarray=event.map((user)=>user.id)
      this.setState({ allpeoplePicker2_User:resultarray ,errorSignoffUsers:false});

    }
    else
    {
      this.setState({ allpeoplePicker2_User:[] ,errorSignoffUsers:true})
    }
 
  }
  public Approverpeoplechange = (event) => {

    
    if(event["length"]>0) 
    {
      var resultarray=event.map((user)=>user.id)
     
      this.setState({ allpeoplePicker_User:resultarray ,errorapproverUsers:false});

    }
    else
    {
      this.setState({ allpeoplePicker_User:[] ,errorapproverUsers:true})
    }
  }
 
  public removeDoc = (e) => {
    var targetelement = parseInt(e.currentTarget.id);
    var filesArray = this.state.filePickerResult;
    var removedFiles=this.state.removedFiles;
    removedFiles=filesArray.filter((key, index) => {
   return index == targetelement
         });

    filesArray = filesArray.filter((key, index) => {
    return index!=targetelement
         });
    this.setState({ filePickerResult: filesArray ,removedFiles:removedFiles });
  }
  private draftForm = (): void => {

    this.state.Topic.trim().length > 0 ? "" : this.setState({ errortopicValue: "Topic is required" });
    this.state.description.trim().length > 0 ? "" : this.setState({ errordescriptionValue: "Description is required" });
    this.state.filePickerResult.length == 0 ? this.setState({ errorfileAttach: "Attachments are required" }) : "";
    this.state.allpeoplePicker2_User.length > 0 ? "" : this.setState({ errorSignoffUsers: true });

    if (this.state.Topic.trim().length > 0 && this.state.description.trim().length > 0 && this.state.filePickerResult.length>0 && this.state.allpeoplePicker2_User.length > 0) {
      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.state.ItemId).update({
        Title: "Tailgate",
        ProcessType: "Tailgate",
        TaskIdentifier: this.state.Topic,
        Description: this.state.description,
        Status: "Draft",
        ApproversId: {
          results: this.state.allpeoplePicker_User.length > 0 ?
            this.state.allpeoplePicker_User : [] 
        },
        SignoffsId: {
          results: this.state.allpeoplePicker2_User.length > 0?this.state.allpeoplePicker2_User:[]  
        },
      
      })
        .then((disID) => {

          sp.web.getFolderByServerRelativeUrl("TaskDocuments").folders.add("TaskDocuments" + '/' + this.state.ItemId).then(result => {
            var tobeRemove=this.state.removedFiles;
            if(tobeRemove.length)
            {
              tobeRemove.map((re)=>{
                sp.web.getFileByServerRelativeUrl(re.files).recycle().then(()=>{
                  console.log("deleted")
                });
              });
            }
             var allUploadFiles=this.state.newFiles;

            if (allUploadFiles.length > 0) {
              this.EachfileUpload(allUploadFiles,result,this.state.ItemId);
              }
              else {
                this.setState({ newFiles: [] });
                this.init();
              }
            
       
            alert("Draft saved Successfully..!");
            this.setState({
              Topic: "",
              description: "",
              filePickerResult: [],
              allpeoplePicker_User: [],
              allpeoplePicker2_User: [],
              EditModel:true
            })
          }); 
        }); 
        
    }
    }

    async EachfileUpload(allUploadFiles,result,newId)
    {
      await allUploadFiles.map((eachfileDetails, index) => {
  
      if (eachfileDetails.files["name"]) {
      result.folder.files.add(eachfileDetails.filename, eachfileDetails.files, true)
          .then((fresult) => {
  
              if (allUploadFiles.length <= index + 1) {
                this.setState({ newFiles: [] });
                this.init();
    
              }
          });
      }
    });
    }

  private SaveForm = (): void => {
    this.state.Topic.trim().length > 0 ? "" : this.setState({ errortopicValue: "Topic is required" });
    this.state.description.trim().length > 0 ? "" : this.setState({ errordescriptionValue: "Description is required" });
    this.state.filePickerResult.length == 0 ? this.setState({ errorfileAttach: "Attachments are required" }) : "";
    this.state.allpeoplePicker2_User.length > 0 ? "" : this.setState({ errorSignoffUsers: true });

    if (this.state.Topic.trim().length > 0 && this.state.description.trim().length > 0 && this.state.filePickerResult.length>0 && this.state.allpeoplePicker2_User.length > 0) {
      let today = new Date().toISOString().slice(0, 10);
      var UserName=this.props.spcontext.pageContext.user.displayName;
      var comments=this.state.description;
      var date =new Date().toLocaleDateString();
      var setSummary=UserName+"~New Request~"+comments+"~"+date+"|";

      sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(this.state.ItemId).update({
        Title: "Tailgate",
        Description: this.state.description,
        Status: "Submit",
        ProcessType: "Tailgate",
        TaskIdentifier: this.state.Topic,
        ApproversId: {
          results: this.state.allpeoplePicker_User.length > 0 ? this.state.allpeoplePicker_User : [] // User/Groups ids as an array of numbers
        },

        SignoffsId: {
          results: this.state.allpeoplePicker2_User.length > 0 ?this.state.allpeoplePicker2_User:[]  // User/Groups ids as an array of numbers
        },
        ApprovalSummary:setSummary
      })
        .then((disID) => { 
          sp.web.getFolderByServerRelativeUrl("TaskDocuments").folders.add("TaskDocuments" + '/' + this.state.ItemId).then(result => {
            var tobeRemove=this.state.removedFiles;
            if(tobeRemove.length)
            {
              tobeRemove.map((re)=>{
                sp.web.getFileByServerRelativeUrl(re.files).recycle().then(()=>{
                  console.log("deleted");
                });
              });
            }
             var allUploadFiles=this.state.newFiles;

            if (allUploadFiles.length > 0) {
              this.EachfileUpload(allUploadFiles,result,this.state.ItemId);
              }
              else {
                this.setState({ newFiles: [] });
                this.init();
              }
            alert(" saved Successfully..!");
            this.setState({
              Topic: "",
              description: "",
              filePickerResult: [],
              allpeoplePicker_User: [],
              allpeoplePicker2_User: [],
              EditModel:true
            })
          }); 
          });
        
   
    }

  }
  public onFirstDataRendered = (params) => {
    params.api.sizeColumnsToFit();
  };
  public render(): React.ReactElement<ITailGateRequestDashboardProps> {

  
const {get_Active_Paged_array}=this.state;
    var StatusSummary=this.state.StatusSummary.split('|');

    const options: IChoiceGroupOption[] = [
      { key: 'A', text: 'Approve' },
      { key: 'B', text: 'Return' }

    ];
    const modelProps = {
      isBlocking: true,
      topOffsetFixed: false,
    };
    const dialogStyles: Partial<IDialogStyles> = {
      main: [
        {
          fontFamily: "Poppins, sans-serif",
          selectors: {
            ".ms-Dialog-title": {
              fontFamily: "Poppins, sans-serif",
            },
            ".ms-Dialog-subText": {
              fontFamily: "Poppins, sans-serif",
            },
          },
        },
      ],
    };
    const columnstyle: Partial<IStackProps> = { 
      tokens: {
        childrenGap: 5,
      },
      styles: {
        root: {
          width: "100%",
        },
      },
    };
    return (
      <div>
      <div className="container">
        <h2 className="heading-2 mt-0">My Tasks</h2>
        <div className={styles.row}>
          <Pivot  linkFormat={PivotLinkFormat.tabs} >
            <PivotItem linkText="Active Tasks" itemKey="1">

            <div className="ag-theme-alpine" style={ {height: 519, width: "100%" } }>
                <AgGridReact
                 onFirstDataRendered={this.onFirstDataRendered}
                 columnDefs={this.ActivecolumnDefs}
                    rowData={this.state.getActiveDataDetails}  pagination={true}  paginationPageSize={10}>
                </AgGridReact>
            </div>
               
            </PivotItem>
            <PivotItem linkText="Draft Requests" itemKey="2">

              
            <div className="ag-theme-alpine" style={ {height: 519, width: "100%" } }>
                <AgGridReact
                onFirstDataRendered={this.onFirstDataRendered}
                 columnDefs={this.DraftcolumnDefs}
                    rowData={this.state.get_draftDetails}  pagination={true}  paginationPageSize={10}>
                </AgGridReact>
            </div>

            </PivotItem>
            <PivotItem linkText="Completed Tasks" itemKey="3">

                <div className="ag-theme-alpine" style={ {height: 519, width: "100%" } }>
                <AgGridReact
                onFirstDataRendered={this.onFirstDataRendered}
                 columnDefs={this.CompletecolumnDefs}
                    rowData={this.state.get_completeDetails}  pagination={true}  paginationPageSize={10}>
                </AgGridReact>
            </div>
               
            </PivotItem>
            <PivotItem linkText="Read Only Tasks" itemKey="4">
               
            <div className="ag-theme-alpine" style={ {height: 519, width: "100%" } }>
                <AgGridReact
                onFirstDataRendered={this.onFirstDataRendered}
                 columnDefs={this.ReadcolumnDefs}
                    rowData={this.state.get_readonlyDetails}  pagination={true}  paginationPageSize={10}>
                </AgGridReact>
            </div>
            </PivotItem>
          </Pivot>
        </div></div>
       {
     <Dialog
       hidden={this.state.ApprovalModal}
       modalProps={modelProps}
       minWidth="600px"
       styles={dialogStyles}
     >
<div className={styles.tailGateRequestDashboard}>
      <div className="container">
        <div className={classnames(styles.row, styles.nopaddingbottom)}>
          <div className={styles.col_3}> 
            <label className={styles.divalign}>Topic </label>
          </div>
          <div className={styles.col_1}>
          :
          </div>
          <div className={styles.col_6}>
            <label>{this.state.Topic}</label>
          </div>
        </div>
        <div className={classnames(styles.row, styles.nopaddingbottom)}>
          <div className={styles.col_3}>
            <label className={styles.divalign}>Description </label>
          </div>
          <div className={styles.col_1}>
          :
          </div>
          <div className={styles.col_6}>
            <label>{this.state.description}</label>
          </div>
        </div>
        <div className={classnames(styles.row, styles.nopaddingbottom)}>
          <div className={styles.col_3}>
            <label className={styles.divalign}>Attachments  </label>
          </div>
          <div className={styles.col_1}>
          :
          </div>
          <div className={styles.col_6}>
            {
              this.state.fileDetails.map((filedet)=>{
                return (  
                  <div>
                <Link href={filedet.files}>{filedet.filename}</Link>

                  </div> 
                )  
              })      
            }
          </div> 
        </div>
        <hr/>  
        <div className={classnames(styles.row, styles.nopaddingbottom)}>
          <div className={styles.col_6}>
            <label className={styles.divalign}>Action History</label>
            </div>
            <div className={styles.col_12}>
            <table className={styles.table}><thead><tr><th>Name</th><th>Action</th><th>Comments</th><th>Date</th></tr></thead ><tbody>
            {
             StatusSummary.map((rowDet)=>{
               if(rowDet)
               {
                rowDet=rowDet.split('~');
                return(
                <tr><td>{rowDet[0]}</td><td>{rowDet[1]}</td><td>{rowDet[2]}</td><td>{rowDet[3]}</td></tr> 
                )
               }

              })
            }
            </tbody></table>
          </div>
      </div>

        </div>
        <hr/> 
        <div className={styles.row} hidden={this.state.btnsReadonly} >
          <div className={styles.col_6}> 
          <label className={styles.divalign}>Action</label>
            <ChoiceGroup defaultSelectedKey="A" options={this.options} onChange={(e,option)=>{  this.setState({ approveStatus: option.text })}} required={true} />
          </div>

        </div>
        <div className={styles.row} hidden={this.state.btnsReadonly}>
          <div className={classnames(styles.col_12)}>
          <label className={styles.divalign}>Comments</label>
            <TextField multiline  required={this.state.approveStatus=="Approve"?false:true} value={this.state.comments} onChanged={newVal => {
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
        </div>
        <div className={styles.textcenter}>
          <span className={styles.buttonspace}> 
           <DefaultButton text="Cancel" onClick={this._alertClicked} />
          </span>
          
          <span className={styles.buttonspace}  hidden={this.state.btnsReadonly}>
            <PrimaryButton text="Submit" onClick={this.SubmitForm} /></span>
        </div>
      </div>
     </Dialog>
  
       }
       {

          <Dialog
          hidden={this.state.SignOffModal}
          modalProps={modelProps}
          minWidth=" 600px"
          styles={dialogStyles}
          >
          <div className={styles.tailGateRequestDashboard}>
          <div className='container'>
          <div className={styles.row}>
            <div className={styles.col_3}>
              <label className={styles.divalign}>Topic</label>
            </div>
            <div className={styles.col_1}>
              :
              </div>
            <div className={styles.col_6}>
              <label>{this.state.Topic}</label>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_3}>
              <label className={styles.divalign}>Description : </label>
            </div>
            <div className={styles.col_1}>
              :
              </div>
            <div className={styles.col_6}>
              <label>{this.state.description}</label>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_3}>
              <label className={styles.divalign}>Attachments : </label>
            </div>
            <div className={styles.col_1}>
              :
              </div>
            <div className={styles.col_6}>
              {
                this.state.filePickerResult.map((filedet)=>{
                  return (
                    <div>
                  <Link href={filedet.files}>{filedet.filename}</Link>

                    </div>
                  )
                })
              }
            </div>
          </div>
          <hr/>  
          <div className={classnames(styles.row, styles.nopaddingbottom)}>
            <div className={styles.col_6}>
              <label className={styles.divalign}>Action History</label>
              </div>
              <div className={styles.col_12}>
              <table className={styles.table}><thead><tr><th>Name</th><th>Action</th><th>Comments</th><th>Date</th></tr></thead ><tbody>
            {
             StatusSummary.map((rowDet)=>{
               if(rowDet)
               {
                rowDet=rowDet.split('~');
                return(
                <tr><td>{rowDet[0]}</td><td>{rowDet[1]}</td><td>{rowDet[2]}</td><td>{rowDet[3]}</td></tr> 
                )
               }

              })
            }
            </tbody></table>
            </div>
          </div>

          </div>
          <hr/>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <Checkbox label="Sign Off"  onChange={(e,option)=>{  this.setState({ chksignOffStatus:option })}} />
            </div>

          </div>
       
          <div className={styles.textcenter}>
            <span className={styles.buttonspace}>
              <DefaultButton text="Cancel" onClick={(e)=>this.setState({SignOffModal:true,Topic:"", description:"", fileDetails:[]})} />
            </span>
            <span  className={styles.buttonspace}  hidden={this.state.btnsReadonly}>
              <PrimaryButton text="Submit" onClick={this.SignoffForm} /></span>
          </div>
          </div>
          </Dialog>
       } 
       {
       <Dialog
          hidden={this.state.EditModel}
          modalProps={modelProps}
          minWidth="600px"
          styles={dialogStyles}
          >
          <div className={styles.tailGateRequestDashboard}>
          <div className={styles.container}>
          <hr></hr>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <TextField label="Topic" required
                value={this.state.Topic}
                onChanged={newVal => {
                  newVal && newVal.length > 0
                    ? this.setState({
                      Topic: newVal,
                      errortopicValue: ""
                    })
                    : this.setState({
                      Topic: newVal,
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
                context={this.props.spcontext}
                titleText=""
                personSelectionLimit={10}
                groupName={""}
                defaultSelectedUsers={this.state.getApprovepeoplePicker_User}
                showtooltip={false}
                required={false}
                disabled={false}
                ensureUser={true}
                onChange={(e) =>this.Approverpeoplechange.call(this,e)}
                
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
                value={this.state.description}
                onChanged={newDesVal => {
                  newDesVal && newDesVal.length > 0
                    ? this.setState({
                      description: newDesVal,
                      errordescriptionValue: ""
                    })
                    : this.setState({
                      description: newDesVal,
                      errordescriptionValue:
                        "Description is required"
                    });
                }}
                multiline rows={3} errorMessage={this.state.errordescriptionValue} />
            </div>
            <div className={styles.col_6}>
              <Label required>Sign offs</Label>
             <PeoplePicker
                context={this.props.spcontext}
                titleText=""
                personSelectionLimit={10}
                groupName={""}
                showtooltip={false}
                disabled={false}
                ensureUser={true}
                defaultSelectedUsers={this.state.getSignOffUser}
                onChange={(e) =>this.SignOffpeoplechange.call(this,e)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
      
              />   
              {this.state.errorSignoffUsers ? <Label className={styles.pickerlabelErrormsg}>Sign Offs is required</Label> : ""}
  
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.col_6}>
              <div>
                <Label required>Attachments</Label>
                <input type="file" multiple accept=".xlsx,.xls,.doc, image/*, .docx,.ppt, .pptx,.txt,.pdf" onChange={this.fileUploadCallback}
                />
                {this.state.errorfileAttach ? <Label className={styles.pickerlabelErrormsg}>Attachment is required</Label> : ""}
              </div>
              <div className={styles.col_6}>
              {
                this.state.filePickerResult.map((filedet,index)=>{
                 
                  return (
                    <div className={styles.attach}> 
                     
                  <Link href={filedet.files}>{filedet.filename}</Link>

                  
                    <IconButton className={styles.btntransparent} iconProps={this.DeleteIcon} onClick={this.removeDoc.bind(this)} id={index.toString()}>  
                    </IconButton><br></br> 
                  </div>
                  
                  )
                })
              }
            </div>
            </div>
          </div>
          <div className={styles.textcenter}>
          <span className={styles.buttonspace}>
              <PrimaryButton className={styles.btnDraft} text="Save as Draft" onClick={this.draftForm} />
            </span>
            <span className={styles.buttonspace}>
              <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.SaveForm} />
            </span>
            <span className={styles.buttonspace}>
              <DefaultButton  text="Cancel" onClick={()=>this.setState({EditModel:true})} />
            </span>
          </div>
        </div>
          </div>
          </Dialog>
       } 
      
    </div>
    );
  }
}
