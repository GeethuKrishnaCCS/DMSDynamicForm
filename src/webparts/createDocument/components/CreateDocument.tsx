import * as React from 'react';
import styles from './CreateDocument.module.scss';
import { ICreateDocumentProps, ICreateDocumentState } from '../interfaces/ICreateDocumentProps';
import { DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IDropdownOption, IIconProps, IPivotStyles, IconButton, Label, Pivot, PivotItem, PrimaryButton, TextField } from '@fluentui/react';
import SimpleReactValidator from 'simple-react-validator';
import { BaseService } from '../services';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import DynamicForms from './DynamicForms';
import CustomFileInput from './CustomFileInput';
export default class CreateDocument extends React.Component<ICreateDocumentProps, ICreateDocumentState, {}> {
  private _service: BaseService;
  private validator: SimpleReactValidator;
  private myfile:any;
  private documentNameExtension: string;
 public constructor(props: ICreateDocumentProps) {
    super(props);
    this.state = {
      currentUserId: null,
      currentUserName: "",
      currentUserEmail: "",
      title: "",
      department: "",
      departmentId: null,
      departmentOption: [],
      departmentCode: "",
      category: "",
      categoryId: null,
      categoryOption: [],
      categoryCode:"",
      contentTypeArray: [],
      contentTypeId: "",
      sourcecontentTypeId: "",
      publishcontentTypeId: "",
      contentTypeName: "",
      sourcelistID: "",
      publishlistID: "",
      listID: "",
      reviewers: [],
      reviewersDetails: [],
      reviewersEmail: "",
      reviewersName: "",
      approver: "",
      approverEmail: "",
      approverName: "",
      disableDynamic: false,
      dynamic:false,
      cancelConfirmMsg: "none",
      confirmDialog: true,
      selectedKey:0,
      proceedButton:false,
      hideEdit: true,
      showSubmitBTN: false,
      hideSubmitBTN: true,
      DocumentId:null,
      mydoc: null as File | null,
      incrementSequenceNumber: "",
      documentid: "", 
      documentName: ""
    }
    this._service = new BaseService(this.props.context, this.props.siteUrl);
    this.bindDropdown = this.bindDropdown.bind(this);
    this.getListGuid = this.getListGuid.bind(this);
    this.departmentChange = this.departmentChange.bind(this);
    this.categoryChange = this.categoryChange.bind(this);
    this._selectedReviewers = this._selectedReviewers.bind(this);
    this._selectedApprover = this._selectedApprover.bind(this);
    this.onProceedClick = this.onProceedClick.bind(this);
    this.submitCallBack = this.submitCallBack.bind(this);
    this.saveCallBack = this.saveCallBack.bind(this);
        this.onNextClick = this.onNextClick.bind(this);
        this.onPreviousClick = this.onPreviousClick.bind(this);

  }
  // Validator
  public UNSAFE_componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: { required: "This field is mandatory" }
    });
  }
  // On load
  public async componentDidMount() {
    const user = await this._service.getCurrentUser();
    console.log(user)
    this.setState({
      currentUserId: user.Id,
      currentUserName: user.Title,
      currentUserEmail: user.Email
    });

    await this.bindDropdown();
  }
  public async bindDropdown() {
    const departmentArray: any = [];
    const departmentListdata = await this._service.getListItems(this.props.siteUrl, this.props.department);
    console.log(departmentListdata);
    for (let i = 0; i < departmentListdata.length; i++) {
      const departmentdata = {
        key: departmentListdata[i].ID,
        text: departmentListdata[i].Title,
      };
      departmentArray.push(departmentdata);
    }
    this.setState({
      departmentOption: departmentArray,
    });
  }
  //ContractName Change
  private titleChange = (ev: React.FormEvent<HTMLInputElement>, Title: string): void => {
    this.setState({
      title: Title,
    });
  }
  //Department Change
  public async departmentChange(event: React.FormEvent<HTMLDivElement>, departmentitem: IDropdownOption) {
    const categoryArray: any = [];
    this.setState({ departmentId: departmentitem.key, department: departmentitem.text, categoryId: "", approverName: "" });
    const departmentdata = await this._service.getItemById(this.props.siteUrl, this.props.department, Number(departmentitem.key))
    this.setState({ departmentCode: departmentdata.Code });
    //Get Contract Type
    const categorydataItems = await this._service.getCategoryItems(this.props.siteUrl, this.props.category, departmentitem.key)
    for (let i = 0; i < categorydataItems.length; i++) {
      const categorydata = {
        key: categorydataItems[i].ID,
        text: categorydataItems[i].Title,
      };
      categoryArray.push(categorydata);
    }
    this.setState({
      contentTypeArray: categorydataItems,
      categoryOption: categoryArray,
    });
  }
  //Category Change
  public categoryChange = async (event: React.FormEvent<HTMLDivElement>, categoryitem: IDropdownOption) => {
    let ApproverId: any = null;
    let ApproverName: string = "";
    this.setState({ contentTypeName: "", categoryId: categoryitem.key, category: categoryitem.text });
    //Get Contract Type
    const selectedType = this.state.contentTypeArray.filter((item) => item.ID === categoryitem.key)
    await this.getListGuid(selectedType[0].ContentTypeName);
    console.log(selectedType)
    if (selectedType[0].Approver.ID !== null) {
      ApproverId = selectedType[0].Approver.ID;
      ApproverName = selectedType[0].Approver.Title;
  }
  this.setState({
    categoryCode: selectedType[0].Code,
    contentTypeName: selectedType[0].ContentTypeName,
    approver: ApproverId,
    approverName: ApproverName,
    dynamic: false,
});
  }
  private getListGuid(ContractContentTypeName: string) {
    this.setState({ sourcecontentTypeId: "", sourcelistID: "",publishcontentTypeId: "", publishlistID: "", disableDynamic: true })
    let sourcereslistid: string = "";
    let publishreslistid: string = "";
    this._service.getListGuid(this.props.siteUrl, this.props.sourceDocument)
      .then(sourceres => {
        sourcereslistid = sourceres.Id;
        this._service.getContentTypeId(this.props.siteUrl, this.props.sourceDocument)
          .then(data => {
            const contentType = data.filter((item: any) => item.Name === ContractContentTypeName);
            if (contentType.length > 0) {
              this.setState({ sourcelistID: sourcereslistid, sourcecontentTypeId: contentType[0].Id.StringValue });
            }
          });
      });
      this._service.getListGuid(this.props.siteUrl, this.props.publishDocument)
      .then(publishres => {
        publishreslistid = publishres.Id;
        this._service.getContentTypeId(this.props.siteUrl, this.props.publishDocument)
          .then(data => {
            const contentType = data.filter((item: any) => item.Name === ContractContentTypeName);
            if (contentType.length > 0) {
              this.setState({ publishlistID: publishreslistid, publishcontentTypeId: contentType[0].Id.StringValue });
            }
          });
      });
  }
  //Reviewer Change
  public _selectedReviewers = (items: any[]) => {
    const reviewerEmail :any = [];
    const getSelectedReviewers :any = [];
    const revieweritems :any = [];

    items.forEach(async (rewitem: any) => {
        await this._service.EnsureUser(rewitem.secondaryText)
            .then((reviewer: any) => {
                reviewerEmail.push(rewitem.secondaryText);
                getSelectedReviewers.push(reviewer.data.Id);
                revieweritems.push({
                    ID: reviewer.data.Id,
                    EMail: rewitem.secondaryText
                });
            })

    });
    this.setState({ reviewers: getSelectedReviewers, reviewersDetails: revieweritems, reviewersEmail: reviewerEmail });

}
//Approver Change
public _selectedApprover = async (items: any[]) => {
  console.log("_selectedApprover ", items)

  let approverEmail;
  let approverName;
  const getSelectedApprover:any = [];
  for (let item = 0; item < items.length; item++) {
      await this._service.EnsureUser(items[item].secondaryText)
          .then((approver: any) => {
              approverEmail = items[item].secondaryText;
              approverName = items[item].text;
              getSelectedApprover.push(approver.data.Id);
          })

  }
  this.setState({ approver: getSelectedApprover[0], approverEmail: approverEmail, approverName: approverName });

}
public add = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const filedata = e.target.files !== null ? e.target.files[0] : "";
  this.myfile = filedata
  // @ts-ignore: Object is possibly 'null'.
  this.setState({ ...this.state, mydoc: this.myfile });
 

}
//On proceed click
public onProceedClick(Source :any) {
  if(this.state.title !== "" && this.state.department !== "" && this.state.category !== ""){
    if(this.state.department === "HR" && this.state.category === "Employees master records"){
      this._documentidgeneration();
    }
  }
  this.setState({ disableDynamic: false, proceedButton :true, selectedKey: (Number(this.state.selectedKey) + 1) % 3 });
}
//Documentid generation
public async _documentidgeneration() {
  let separator;
  let sequenceNumber;
  let idcode;
  let counter;
  var incrementstring;
  let increment;
  let documentid;
  let isValue = "false";
  let settingsid;
  let documentname;
  // Get document id settings
  const documentIdSettings = await this._service.getListItems(
    this.props.siteUrl,
    this.props.documentIdSettings
  );
  console.log("documentIdSettings", documentIdSettings);
  separator = documentIdSettings[0].Separator;
  sequenceNumber = documentIdSettings[0].SequenceDigit;
  idcode = this.state.departmentCode + separator + this.state.categoryCode;
  if (documentIdSettings) {
    // Get sequence of id
    const documentIdSequenceSettings = await this._service.getListItems(
      this.props.siteUrl,
      this.props.documentIdSequenceSettings
    );
    console.log("documentIdSequenceSettings", documentIdSequenceSettings);
    for (var k in documentIdSequenceSettings) {
      if (documentIdSequenceSettings[k].Title === idcode) {
        counter = documentIdSequenceSettings[k].Sequence;
        settingsid = documentIdSequenceSettings[k].ID;
        isValue = "true";
      }
    }
    if (documentIdSequenceSettings) {
      // No sequence
      if (isValue === "false") {
        increment = 1;
        incrementstring = increment.toString();
        let idsettings = {
          Title: idcode,
          Sequence: incrementstring,
        };
        const addidseq = await this._service.createNewListItem(
          this.props.siteUrl,
          this.props.documentIdSequenceSettings,
          idsettings
        );
        if (addidseq) {
          await this._incrementSequenceNumber(
            incrementstring,
            sequenceNumber
          );

          if (this.state.departmentCode !== "") {
            documentid =
              this.state.departmentCode +
              separator +
              this.state.categoryCode +
              separator +
              this.state.incrementSequenceNumber;
          } else {
            documentid =
              this.state.departmentCode +
              separator +
              this.state.categoryCode +
              separator +
              this.state.incrementSequenceNumber;
          }
          documentname = documentid + " " + this.state.title;

          this.setState({
            documentid: documentid,
            documentName: documentname,
          });
          await this.documentCreation();
        }
      }
      // Has sequence
      else {
        increment = parseInt(counter) + 1;
        incrementstring = increment.toString();
        let idItems = {
          Title: idcode,
          Sequence: incrementstring,
        };
        // const afterCounter = await this._Service.itemUpdate(this.props.siteUrl, this.props.documentIdSequenceSettings, settingsid, idItems);
        const afterCounter = await this._service.updateItem(this.props.siteUrl,this.props.documentIdSequenceSettings,settingsid,idItems);
        if (afterCounter !== undefined) {
          await this._incrementSequenceNumber(incrementstring,sequenceNumber);
          if (this.state.departmentCode !== "") {
            documentid =
              this.state.departmentCode +
              separator +
              this.state.categoryCode +
              separator +
              this.state.incrementSequenceNumber;
          } else {
            documentid =
              this.state.departmentCode +
              separator +
              this.state.categoryCode +
              separator +
              this.state.incrementSequenceNumber;
          }
          documentname = documentid + " " + this.state.title;
          this.setState({
            documentid: documentid,
            documentName: documentname,
          });
          await this.documentCreation();
        }
      }
    }
  }
}
// Append sequence to the count
public _incrementSequenceNumber(incrementvalue: string, sequenceNumber: number) {
  let incrementSequenceNumber = incrementvalue;
  while (incrementSequenceNumber.length < sequenceNumber)
    incrementSequenceNumber = "0" + incrementSequenceNumber;
  console.log(incrementSequenceNumber);
  this.setState({
    incrementSequenceNumber: incrementSequenceNumber,
  });
}
 // Add Source Document metadata
public async documentCreation(){
  let documentNameExtension;
    let sourceDocumentId;
    let docinsertname;
  if (this.state.mydoc !== null) {
   const myfile = this.state.mydoc;
    var splitted = this.state.mydoc.name.split(".");
    documentNameExtension = this.state.documentName + "." + splitted[splitted.length - 1];
    this.documentNameExtension = documentNameExtension;
    docinsertname =
      this.state.documentid + "." + splitted[splitted.length - 1];
    if (myfile.size) {
      // add file to source library
      const fileUploaded = await this._service.uploadDocument(
        docinsertname,
        myfile,
        this.props.sourceDocument
      );
      if (fileUploaded) {
        const filePath =
          window.location.protocol +
          "//" +
          window.location.host +
          fileUploaded.data.ServerRelativeUrl;
          if(filePath){
        const item = await fileUploaded.file.getItem();
        console.log(item);
        sourceDocumentId = item["ID"];
        this.setState({ DocumentId: sourceDocumentId });
        // update metadata
         // Without Expiry Date
    // let WorkflowStatus: string;
    // let Workflow: string;
    // if (this.state.reviewers.length !== 0) {
    //   WorkflowStatus = "Under Review";
    //   Workflow = "Review";
    // } else {
    //   WorkflowStatus = "Under Approval";
    //   Workflow = "Approval";
    // }
    let sourceUpdate = {
        Title: this.state.title,
        DocumentID: this.state.documentid,
        DocumentName: this.documentNameExtension,
        Category: this.state.category,
        ApproverId: this.state.approver,
        ReviewersId: this.state.reviewers,
        OwnerId: this.state.currentUserId,
        Revision: "0",
        WorkflowStatus: "Draft",
        DocumentStatus: "Active",
        Workflow:  "Draft",
       DepartmentName: this.state.department,
        
      };
      await this._service.updateLibraryItem(
        this.props.siteUrl,
        this.props.sourceDocument,
        this.state.DocumentId,
        sourceUpdate
      );
        // await this.addSourceDocument();
        if (item) {
          // await this._triggerPermission(sourceDocumentId);
         
        }
      }
      }
    }
  }
   
    
    
  
}
public savepublish(){

  const dataforDocument = {
    DMSTitle : this.state.title,
    DepartmentId : this.state.departmentId,
    CategoryId : this.state.categoryId
  }
  this._service.createNewLibraryItem(this.props.siteUrl,this.props.publishDocument,dataforDocument)
  .then((response: any) => {
    if (response.data !== null) {
      this.setState({DocumentId : response.data.ID,listID : this.state.publishlistID , contentTypeId : this.state.publishcontentTypeId})
    }

  });
}
public submitCallBack(callbackdata: any) {
  this.setState({
      hideSubmitBTN: callbackdata.disableSubmitBTN,
      disableDynamic: callbackdata.disableDynamic,
      showSubmitBTN: callbackdata.showSubmitBTN,
      hideEdit: false
  })
}
public async saveCallBack(callbackdata: any) {
 
      this.setState({
          showSubmitBTN: callbackdata.showSubmitBTN,
          disableDynamic: callbackdata.disableDynamic,
          hideEdit: false,
          selectedKey: (Number(this.state.selectedKey) + 1) % 3
      });
  
}
public onNextClick() {
  this.setState({ selectedKey: (Number(this.state.selectedKey) + 1) % 3 });
}
public async onPreviousClick() {
 this.setState({ selectedKey: (Number(this.state.selectedKey) - 1) % 3 });
 
}
//Cancel Document
private onCancel = () => {
  this.setState({
    cancelConfirmMsg: "",
    confirmDialog: false,
  });
}
private dialogContentProps = {
  type: DialogType.normal,
  closeButtonAriaLabel: 'none',
  title: 'Do you want to cancel?',
  //subText: '<b>Do you want to cancel? </b> ',
};
private dialogStyles = { main: { maxWidth: 500 } };
private modalProps = {
  isBlocking: true,
};
//For dialog box of cancel
private _dialogCloseButton = () => {
  this.setState({
    cancelConfirmMsg: "none",
    confirmDialog: true,
  });
}
//Cancel confirm
private _confirmYesCancel = () => {
  this.setState({
    cancelConfirmMsg: "none",
    confirmDialog: true,
  });
  window.location.replace(this.props.siteUrl);
}
//Not Cancel
private _confirmNoCancel = () => {
  this.setState({
    cancelConfirmMsg: "none",
    confirmDialog: true,
  });
}
  public render(): React.ReactElement<ICreateDocumentProps> {
    const pivotStyles: IPivotStyles = {
      root: {
          padding: '10px'
      },
      link: '',
      linkIsSelected: '',
      linkContent: '',
      text: '',
      count: '',
      icon: '',
      linkInMenu: '',
      overflowMenuButton: ''
  }
  const NextIcon: IIconProps = {
    iconName: 'PageRight', styles: {
        root: {
            color: 'rgb(7, 27, 117)',
            backgroundColor: 'white',
            fontSize: '2.3em',
            padding: '2px'
        }
    }
};
const PrevIcon: IIconProps = {
    iconName: 'PageLeft', styles: {
        root: {
            color: 'rgb(7, 27, 117)',
            backgroundColor: 'white',
            fontSize: '2.3em',
            padding: '2px'
        }
    }
};
    return (
      <section className={`${styles.createDocument}`}>
        <div className={styles.border}>
          <div className={styles.documenttitle}>{this.props.webpartHeader}</div>
          <div className={styles.formsection}>
          <Pivot styles={pivotStyles} aria-label="Links of Tab Style Pivot Example" selectedKey={String(this.state.selectedKey)} linkFormat="tabs" style={{ paddingLeft: '15px' }}>
                        <PivotItem headerText="Document Info" style={{ paddingRight: "2em" }}>
                        <div >
              <TextField label="Title" id="t1"
                onChange={this.titleChange}
                value={this.state.title} required />
              <div style={{ color: "#dc3545" }}>{this.validator.message("Title", this.state.title, "required|alpha_num_dash_space|max:200")}{" "}</div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Dropdown label='Department' required
                  selectedKey={this.state.departmentId}
                  placeholder="Select an option"
                  options={this.state.departmentOption}
                  onChange={this.departmentChange}
                />
                <div style={{ color: "#dc3545" }}>
                  {this.validator.message("department", this.state.departmentId, "required")}{" "}
                </div>
              </div>
              <div className={styles.wdthlft}>
                <Dropdown label='Category'
                  selectedKey={this.state.categoryId}
                  placeholder="Select an option"
                  options={this.state.categoryOption}
                  onChange={this.categoryChange}
                />

              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", }} >
                  <CustomFileInput
                    onChange={this.add}
                    key={1} />
                  <div ><Label className={styles.cmnlabel}>
                    {this.myfile !== undefined ? this.myfile.name : ""}
                  </Label>
                  </div>
                </div>
            {this.state.category !== "Employees master records" && this.state.category !== "" && this.state.department === "HR"   &&<div>
            <div className={styles.divrow}>
                                    <div className={styles.wdthrgt} >
                                        <PeoplePicker
                                            context={this.props.context as any}
                                            titleText="Reviewer(s)"
                                            personSelectionLimit={3}
                                            groupName={""} // Leave this blank in case you want to filter from all users
                                            showtooltip={true}
                                            disabled={false}
                                            ensureUser={true}
                                            showHiddenInUI={false}
                                            onChange={this._selectedReviewers}
                                            // selectedItems={(items) => this._selectedReviewers(items)}
                                            defaultSelectedUsers={this.state.reviewersName}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000} />
                                        
                                    </div>
                                    <div className={styles.wdthlft}>
                                        <PeoplePicker
                                            context={this.props.context as any}
                                            titleText="Approver *"
                                            personSelectionLimit={1}
                                            groupName={""} // Leave this blank in case you want to filter from all users    
                                            showtooltip={true}
                                            disabled={false}
                                            ensureUser={true}
                                            onChange={this._selectedApprover}
                                            showHiddenInUI={false}
                                            defaultSelectedUsers={[this.state.approverName]}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={1000} />
                                        <div style={{ color: "#dc3545" }}>
                                            {this.validator.message("approverName", this.state.approverName, "required")}{" "}
                                        </div>
                                    </div>
                                </div>
                                </div>}
                                {this.state.proceedButton === false &&<div className={styles.rgtalign} >
                                   <PrimaryButton id="b2" className={styles.btn} onClick={() => this.onProceedClick("ProceedButton")}>Proceed</PrimaryButton > 
                                    <PrimaryButton id="b3" className={styles.btn} onClick={this.onCancel}>Cancel</PrimaryButton >
                                </div>}
                                {this.state.proceedButton === true && <div style={{ width: "100%", textAlign: "right", paddingTop: "2em" }}>
                                    <IconButton iconProps={NextIcon} title='Next' onClick={this.onNextClick} />
                                </div>}
 {/* {/ Cancel Dialog Box /} */}
 <div style={{ display: this.state.cancelConfirmMsg }}>
                <div>
                  <Dialog
                    hidden={this.state.confirmDialog}
                    dialogContentProps={this.dialogContentProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>
              </div>
                          </PivotItem>
                          <PivotItem headerText="Additional Info" style={{ paddingRight: "1em" }}>
                            {this.state.hideEdit === true && <div style={{ width: "100%", textAlign: "right", paddingTop: "2em" }}>
                                <IconButton iconProps={PrevIcon} title='Previous' onClick={this.onPreviousClick} style={{ marginRight: "1em" }} />
                                <IconButton iconProps={NextIcon} title='Next' onClick={this.onNextClick} style={{ marginRight: "1em" }} />
                            </div>}
                            {this.state.dynamic === false &&
                                <DynamicForms context={this.props.context}
                                    siteUrl={this.props.siteUrl}
                                    contractIndex={this.props.sourceDocument}
                                    contentTypeName={this.state.contentTypeName}
                                    submitCallBack={this.submitCallBack}
                                    contractIndexId={this.state.DocumentId}
                                    listID={this.state.listID}
                                    contentTypeId={this.state.contentTypeId}
                                    disableDynamic={this.state.disableDynamic}
                                    saveCallBack={this.saveCallBack}
                                    hideEdit={this.state.hideEdit}
                                    absolutesiteUrl={this.props.absolutesiteUrl} />
                            }
                            {/* {/ Cancel Dialog Box /} */}
                            <div style={{ display: this.state.cancelConfirmMsg }}>
                                <div>
                                    <Dialog
                                        hidden={this.state.confirmDialog}
                                        dialogContentProps={this.dialogContentProps}
                                        onDismiss={this._dialogCloseButton}
                                        styles={this.dialogStyles}
                                        modalProps={this.modalProps}>
                                        <DialogFooter>
                                            <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                                            <DefaultButton onClick={this._confirmNoCancel} text="No" />
                                        </DialogFooter>
                                    </Dialog>
                                </div>
                            </div>
                            {this.state.hideEdit === false && <div style={{ width: "100%", textAlign: "right", paddingTop: "2em" }}>
                                <IconButton iconProps={PrevIcon} title='Previous' onClick={this.onPreviousClick} style={{ marginRight: "1em" }} />
                                <IconButton iconProps={NextIcon} title='Next' onClick={this.onNextClick} style={{ marginRight: "1em" }} />
                            </div>}
                        </PivotItem>
                          </Pivot>
            
          
          </div>
        </div>
      </section>
    );
  }
}
