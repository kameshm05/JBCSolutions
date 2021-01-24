export interface IReportSignOffState {
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
    isChecked:boolean;
    isreadOnly:boolean;
    isSignOffsummary:string;
    createdBy:string;
    CheckedSignOffArray:any[];
    AllStatusSummary:string;
    signOffCount:string;
    ApproversList:string,
    SignOffList:string,
    Requester:string;
    RequestMail:string;
    RequesterJobTitle:string;
    RequestDate:string;
    RequesterDept:string;
  }
  export interface fileDetails {
    filename:string;
   files:string;
 }