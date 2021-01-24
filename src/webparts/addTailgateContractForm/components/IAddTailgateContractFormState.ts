export interface IAddTailgateContractFormState {
    contractOptions: string;
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
    removedFiles:any[];
    newFiles:any[];
    TypeOfContract:string;
    currentUserId:string;
    TypeOfContractId:string;
    hideAlert:boolean;
    errorfilesizeMsg:string;
    HSEProcessType:string;
    PurchasingProcessType:string;
    FinanceProcessType:string;
    AttachmentType: any[],
    errorAttachmentType: string;
    SelectedAttachementType:string;
    SelectedAttachementTypeID:number;
  }