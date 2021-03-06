export interface IContractorApproveHseState {
    approveOptions: string;
    approvepurchaseOptions:string;
    revalidationOptions:any[];
    revalidationOptionsText:string;
    revalidationOptionsID:number;
    contractorNumber:string;
    erroecontractorNumber:string;
    contractorName:string;
    erroecontractorName:string;
    classification:any[];
    classificationText:string;
    classificationID:number;
    errorclassification:string;
    validationPeriod:number;
    errorvalidationPeriod:string;
    filePickerResult:any[];
    errorfileAttach:string;
    btnSubmitText:string;
    queryStringId:string;
    currentUserId:string;
    isHSEGroupApprover:boolean;
    isPurchasingGroupApprover:boolean;
    isFinanceGroupApprover:boolean;
    hideHseApprove:boolean;
    hidePurchasingApprove:boolean;
    hideFinanceApprove:boolean;
    isHseStatus:string;
    isHseLevel:string;
    isHseDate:string;
    isPurchasingStatus:string;
    isPurchasingLevel:string;
    isPurchasingDate:string;
    axNumber:string;
    erroraxNumber:string;
    isnotApprover:boolean;
    LevelSummary:string;
    HSEStatus:string,
    PurchasingStatus:string,
    FinanceStatus:string;
    endUserHide:boolean;
    hidereturnBox:boolean;
    returncomments:string;
    errorreturncomments:string;
    hideAlert:boolean;
    HSEProcessType:string;
    PurchasingProcessType:string;
    FinanceProcessType:string;
    HSENeededCount:number;
    HSEApprovedCount:number;
    PurchasingNeededCount:number;
    PurchasingApprovedCount:number;
    TypeOfContract:string;
    FinanceNeededCount:number;
    FinanceApprovedCount:number;
    AttachmentType:string;
  }