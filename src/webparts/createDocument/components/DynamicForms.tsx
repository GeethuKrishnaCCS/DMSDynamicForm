import * as React from 'react';
import styles from './DynamicForms.module.scss';
import { IDynamicFormsProps, IDynamicFormsState } from '../interfaces/IDynamicForms';
//import { DynamicFormsServices } from '../services/DynamicFormsService';
import { DynamicForm } from '@pnp/spfx-controls-react/lib/DynamicForm';
import { DefaultButton, Dialog, DialogFooter, DialogType, IIconProps, IconButton, Label, PrimaryButton } from '@fluentui/react';
import { IDynamicFieldProps } from '@pnp/spfx-controls-react/lib/controls/dynamicForm/dynamicField';
import { DynamicFormsServices } from '../services/DynamicFormsService';
import Toast from '../shared/controls/Toast';

export default class DynamicForms extends React.Component<IDynamicFormsProps, IDynamicFormsState> {
    private _dfservice: DynamicFormsServices;
    public constructor(props: IDynamicFormsProps) {
        super(props);
        this.state = {
            listID: "",
            contentTypeId: "",
            itemId: null,
            ChangedTitle: "",
            disableDynamic: false,
            hideEdit: true,
            cancelConfirmMsg: "none",
            confirmDialog: true,
        }
        this._dfservice = new DynamicFormsServices(this.props.context, this.props.absolutesiteUrl);
        this.onEditClick = this.onEditClick.bind(this);
        this.OnCancel = this.OnCancel.bind(this);
    }
    public componentDidMount() {
        this.setState({
            contentTypeId: this.props.contentTypeId,
            listID: this.props.listID,
            itemId: this.props.contractIndexId,
            disableDynamic: this.props.disableDynamic,
            hideEdit: this.props.hideEdit
        });
    }
    public componentDidUpdate(prevProps: Readonly<IDynamicFormsProps>, prevState: Readonly<IDynamicFormsState>, snapshot?: any): void {
        if (this.props !== prevProps) {
            this.setState({ contentTypeId: "", listID: "", itemId: null })
            this.setState({
                contentTypeId: this.props.contentTypeId,
                listID: this.props.listID,
                itemId: Number(this.props.contractIndexId),
                disableDynamic: this.props.disableDynamic,
                hideEdit: this.props.hideEdit
            });
        }

    }

    private renderTitle = (fieldProperties: IDynamicFieldProps): React.ReactElement<IDynamicFieldProps> => {
        return <div >
            <Label> </Label>
        </div>
    }

    public onEditClick() {
        this.setState({ disableDynamic: false });
        const callbackdata = {
            disableSubmitBTN: true,
            disableDynamic: false,
            showSubmitBTN: false
        }

        this.props.submitCallBack(callbackdata);

    }
    //Cancel Form
    public async OnCancel() {
        this.setState({
            cancelConfirmMsg: "",
            confirmDialog: false,
        });
    }
    private _confirmYesCancel = async () => {
        const mandatory: any[] = await this._dfservice._getMandatory(this.props.siteUrl, this.props.contractIndex,
            this.state.contentTypeId);
        console.log(mandatory)
        if (mandatory.length > 0) {
            this.setState({
                disableDynamic: true,
                cancelConfirmMsg: "none",
                confirmDialog: true,
                hideEdit: false
            });
            const callbackdata = {
                disableSubmitBTN: false,
                disableDynamic: true,
                showSubmitBTN: false,
                hideEdit: false
            }

            this.props.submitCallBack(callbackdata);
        }
        else {
            this.setState({
                disableDynamic: true,
                hideEdit: false,
                cancelConfirmMsg: "none",
                confirmDialog: true
            });
            const callbackdata = {
                disableSubmitBTN: false,
                disableDynamic: true,
                showSubmitBTN: true,
                hideEdit: false
            }

            this.props.submitCallBack(callbackdata);
        }
    }
    //Not Cancel
    private _confirmNoCancel = () => {
        this.setState({
            cancelConfirmMsg: "none",
            confirmDialog: true,
        });
    }
    //For dialog box of cancel
    private _dialogCloseButton = () => {
        this.setState({
            cancelConfirmMsg: "none",
            confirmDialog: true,
        });
    }
    private dialogStyles = { main: { maxWidth: 500 } };
    private dialogContentProps = {
        type: DialogType.normal,
        closeButtonAriaLabel: 'none',
        title: 'Do you want to cancel?'
    };
    private modalProps = {
        isBlocking: true,
    };
    public render(): React.ReactElement<IDynamicFormsProps> {
        const EditIcon: IIconProps = { iconName: 'Edit' };
        return (
            <section className={`${styles.dynamicform}`}>
                {this.state.hideEdit === false && <div style={{ width: "100%", textAlign: "right" }}>
                    <Label>Edit <IconButton iconProps={EditIcon} onClick={this.onEditClick} /></Label>
                </div>}
                {this.state.disableDynamic === false && <div>
                    {(this.state.listID !== "" && this.state.contentTypeId !== "") &&
                        <>
                            <DynamicForm
                                context={this.props.context as any}
                                webAbsoluteUrl={this.props.absolutesiteUrl}
                                listId={this.state.listID}
                                listItemId={this.state.itemId}
                                disabled={false}
                                contentTypeId={this.state.contentTypeId}
                                onCancelled={this.OnCancel}
                                onBeforeSubmit={async (listItem) => { return false; }}
                                onSubmitError={(listItem, error) => { console.log(error.message);Toast("warning", "Invalid field value"); }}
                                fieldOverrides={{ "Title": this.renderTitle }}
                                onSubmitted={async (listItemData) => {
                                    const callbackdata = {
                                        showSubmitBTN: true,
                                        disableDynamic: true
                                    }
                                    this.props.saveCallBack(callbackdata);
                                    this.setState({ disableDynamic: true, hideEdit: false });
                                }} />

                        </>}
                </div>}
                {this.state.disableDynamic === true && <div>
                    {(this.state.listID !== "" && this.state.contentTypeId !== "") &&
                        <>
                            <DynamicForm
                                context={this.props.context as any}
                                webAbsoluteUrl={this.props.absolutesiteUrl}
                                listId={this.state.listID}
                                listItemId={this.state.itemId}
                                disabled={true}
                                contentTypeId={this.state.contentTypeId}
                                onCancelled={this.OnCancel}
                                onBeforeSubmit={async (listItem) => { return false; }}
                                onSubmitError={(listItem, error) => { alert(error.message); }}
                                fieldOverrides={{ "Title": this.renderTitle }}
                                onSubmitted={async (listItemData) => {
                                    const callbackdata = {
                                        showSubmitBTN: true
                                    }
                                    this.props.saveCallBack(callbackdata);
                                    this.setState({ disableDynamic: true, hideEdit: false });
                                }} />

                        </>}
                </div>}

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
            </section>
        );
    }
}