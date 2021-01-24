export interface ITailGateRequestDashboardState {
    getActiveDataDetails: any,
    get_draftDetails: any,
    get_completeDetails: any,
    get_readonlyDetails: any,
    get_Active_Paged_array:any;
    get_Draft_Paged_array:any;
    get_Completed_Paged_array:any;
    get_Read_Paged_array:any;
    filterTaskDetails: string,
    filter_draftDetails: string,
    filter_completeDetails: string,
    filter_readonlyDetails: string,
    isTaskView: true,
    isEditView: false,
    Topic: string,
    description: string,
    fileDetails:fileDetails[],
    ApprovalModal:boolean;
    SignOffModal:boolean;
    ItemId:number;
    approveStatus:string;
    chksignOffStatus:boolean;
    comments:string;
    errorcomments:string;
    StatusSummary:string;
    fetchApprovers:any;
    fetchSignOffUsers:any;
    btnsReadonly:boolean;
    EditModel:boolean;
    errortopicValue: string;
    errordescriptionValue: string;
    filePickerResult:any;  
    getUsers:any,
    allpeoplePicker_User:any
    getApprovepeoplePicker_User:any
    getSignOffUser:any;
    topicValue:string;
    descriptionValue:string;
    errorSignoffUsers:boolean;
    errorapproverUsers:boolean;
    allpeoplePicker2_User:any;
    errorfileAttach:string;
    removedFiles:any;
    newFiles:any;
    getoldfilePickerResult:any;
    ActivePageDetails:any;
    isAdmin:boolean;
  }
  export interface fileDetails {
    filenname:string;
   files:string;
 }