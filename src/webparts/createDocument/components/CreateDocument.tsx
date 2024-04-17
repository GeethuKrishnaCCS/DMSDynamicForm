import * as React from 'react';
import styles from './CreateDocument.module.scss';
import { ICreateDocumentProps, ICreateDocumentState } from '../interfaces/ICreateDocumentProps';
import { Dropdown, IDropdownOption, PrimaryButton, TextField } from '@fluentui/react';
import SimpleReactValidator from 'simple-react-validator';
import { BaseService } from '../services';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
export default class CreateDocument extends React.Component<ICreateDocumentProps, ICreateDocumentState, {}> {
  private _service: BaseService;
  private validator: SimpleReactValidator;
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
      contentTypeName: "",
      listID: "",
      reviewers: [],
      reviewersDetails: [],
      reviewersEmail: "",
      reviewersName: "",
      approver: "",
      approverEmail: "",
      approverName: "",
      disableDynamic: false,
      dynamic:false
    }
    this._service = new BaseService(this.props.context, this.props.siteUrl);
    this.bindDropdown = this.bindDropdown.bind(this);
    this.departmentChange = this.departmentChange.bind(this);
    this.categoryChange = this.categoryChange.bind(this);
    this._selectedReviewers = this._selectedReviewers.bind(this);
    this._selectedApprover = this._selectedApprover.bind(this);
    

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
    this.setState({ contentTypeId: "", listID: "", disableDynamic: true })
    let reslistid: string = "";
    this._service.getListGuid(this.props.siteUrl, this.props.sourceDocument)
      .then(res => {
        reslistid = res.Id;
        this._service.getContentTypeId(this.props.siteUrl, this.props.sourceDocument)
          .then(data => {
            const contentType = data.filter((item: any) => item.Name === ContractContentTypeName);
            if (contentType.length > 0) {
              this.setState({ listID: reslistid, contentTypeId: contentType[0].Id.StringValue });
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
  public render(): React.ReactElement<ICreateDocumentProps> {
    return (
      <section className={`${styles.createDocument}`}>
        <div className={styles.border}>
          <div className={styles.documenttitle}>{this.props.webpartHeader}</div>
          <div className={styles.formsection}>
            <div className={styles.divrow}>
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
            {this.state.department === "HR" && this.state.category === "Employees master records" &&<div>
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
                                <div className={styles.rgtalign} >
                                    {/* <PrimaryButton id="b2" className={styles.btn} onClick={() => this.onProceedClick("ProceedButton")}>Proceed</PrimaryButton > */}
                                    {/* <PrimaryButton id="b3" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton > */}
                                </div>
          </div>
        </div>
      </section>
    );
  }
}
