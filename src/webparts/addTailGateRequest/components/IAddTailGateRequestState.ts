export interface IAddTailGateRequestState {
    topicValue: string;
    descriptionValue: string;
    allpeoplePicker_User: any;
    allpeoplePicker2_User: any;
    approverUsers: any;
    SignoffUsers: any;
    errortopicValue: string;
    errordescriptionValue: string;
    errorfileAttach: any;
    errorapproverUsers: boolean;
    errorSignoffUsers: boolean;
    filePickerResult:any;  
    getApprovepeoplePicker_User:any[];
    getSignOffUser:any[];
    addformHide:boolean;
    headerContent:string;
    viewMode:boolean;
    StatusSummary:string;
    fetchSignOffUsers:any;
    removedFiles:any[];
    newFiles:any[]
    fileDetails:fileDetails[];
    approveStatus:string;
    comments:string;
    errorcomments:string;
    fetchApprovers:[];
    StatusCheck:string;
    chksignOffStatus:boolean;
    hideAlert:boolean;
    subText:string;
    singOffUserText:any[];
    approversUserText:any[];
    signoffSummary:string;
    signOffCount:string;
    AssignedDate:string;
    ApproversList:string;
    SignOffList:string;
    pendingAtStage:string;
    ApproverAdminSummary:string;
    adminSelct:boolean;
    adminSelctApprover:string;
    errorDropDown:string;
    isGroupMember:boolean;
    adminApproverid:string;
    Requester:string,
    RequestMail:string,
    RequesterJobTitle:string,
    RequestDate:string,
    RequesterDept:string;
    isModalOpen:boolean;
    hideModal:boolean;
    modalApprovalContent:string;
    signOffModalContent:string;
    hideApproverModal:boolean;
    isApproverModalOpen:boolean;
     errorApproverTxt:string,
    errorSignOffTxt:string,
  }
  export interface fileDetails {
    filename:string;
   files:string;
 }