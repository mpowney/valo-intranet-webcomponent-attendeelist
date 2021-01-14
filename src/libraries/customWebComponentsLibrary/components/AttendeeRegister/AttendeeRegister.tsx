import * as React from 'react';
import { sp } from '@pnp/sp';
import { PermissionKind } from '@pnp/sp/security';
import { ContextualMenu, IContextualMenuItem, IContextualMenuProps } from 'office-ui-fabric-react/lib/ContextualMenu';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { IDragOptions, Modal } from 'office-ui-fabric-react/lib/Modal';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields/list";
import "@pnp/sp/views";
import "@pnp/sp/site-users/web";

import { ProvisionLists } from './ProvisionLists';

export interface IAttendeeRegisterProps {
    _ctx: any;
    calendarIdentifier: string;
    onDemandUrl: string;
    playButtonText: string;
    playButtonIcon: string;
}

export interface IAttendeeRegisterState {
    attendanceData: any[];
    attendanceTypesData: any[];
    showProvisionLists: boolean;
    includeModal: boolean;
    buttonDisabled: boolean;
}

const viewName: string = "All Items";
const titleField: string = "Title";
const createdByField: string = "Author";
const createdByIdField: string = "AuthorId";
const attendanceListTitle: string = "SessionAttendance";
const typesListTitle: string = "AttendanceTypes";
const customAttendanceField: string = "AttendanceType";
const customAttendanceDisplayNameField: string = "AttendanceDisplayName";
const customSortOrderField: string = "SortOrder";
const customIconField: string = "Icon";
const lookupPromiseObjectName: string = "attendeeRegisterAttendancePromise";
const lookupTypesPromiseObjectName: string = "attendeeRegisterAttendanceTypesPromise";
const provisionListModalAlreadyLoadedObjectName: string = "attendeeRegisterProvisionListModalLoaded";

export class AttendeeRegister extends React.Component<IAttendeeRegisterProps, IAttendeeRegisterState> {

    constructor(props: IAttendeeRegisterProps) {
        super(props);

        this.hideProvisionLists = this.hideProvisionLists.bind(this);

        sp.setup({spfxContext: props._ctx});

        this.state = {
            attendanceData: [],
            attendanceTypesData: [],
            showProvisionLists: false,
            includeModal: false,
            buttonDisabled: false
        };
      
    }

    public async componentDidMount() {
        await this.loadAttendance();
        await this.loadAttendanceTypes();

        if (!(window as any)[provisionListModalAlreadyLoadedObjectName]) {
            (window as any)[provisionListModalAlreadyLoadedObjectName]++;
            if ((window as any)[provisionListModalAlreadyLoadedObjectName] == 1) this.setState({ includeModal: true });
        }

    }
    private resetAttendance() {
        delete (window as any)[lookupPromiseObjectName];
    }

    private async loadAttendance() {
        
        this.setState({ buttonDisabled: true });
        const currentUser = await sp.web.currentUser();
        if (!currentUser) {
            console.error(`AttendeeRegister currentUser cannot be found`);
            return;
        }

        try {

            (window as any)[lookupPromiseObjectName] = (window as any)[lookupPromiseObjectName] || sp.web.lists.getByTitle(attendanceListTitle).items.filter(`${createdByIdField} eq ${currentUser.Id}`).getAll();
            const attendanceData = await (window as any)[lookupPromiseObjectName];
            if (!attendanceData) {
                console.error(`attendanceData for user id ${currentUser.Id} (${currentUser.LoginName}) is undefined`);
                return;
            }
            // console.log(attendanceData.filter(item => { return item.Title === this.props.calendarIdentifier; }));
            this.setState({ attendanceData: attendanceData.filter(item => { return item.Title === this.props.calendarIdentifier; }) });

            this.setState({ buttonDisabled: false });

        } catch (err) {

            console.error(`Error returned requesting attendance data:`);
            console.error(err);

            if (await sp.web.currentUserHasPermissions(PermissionKind.ManageLists)) {

                console.error(`Current user has manage lists permission, showing provison lists UI`);
                this.setState({ showProvisionLists: true });

            }

        }

    }

    private async loadAttendanceTypes() {

        this.setState({ buttonDisabled: true });
        try {

            (window as any)[lookupTypesPromiseObjectName] = (window as any)[lookupTypesPromiseObjectName] || sp.web.lists.getByTitle(typesListTitle).items.orderBy(customSortOrderField).getAll();
            const attendanceTypesData = await (window as any)[lookupTypesPromiseObjectName];
            if (!attendanceTypesData) {
                console.error(`attendanceTypesData is undefined`);
                return;
            }
            this.setState({ attendanceTypesData: attendanceTypesData });
            this.setState({ buttonDisabled: false });

        } catch (err) {

            console.error(`Error returned requesting attendance data:`);
            console.error(err);

            if (await sp.web.currentUserHasPermissions(PermissionKind.ManageLists)) {

                console.error(`Current user has manage lists permission, showing provison lists UI: ${typesListTitle}`);
                this.setState({ showProvisionLists: true });

            }

        }

    }

    private hideProvisionLists() {
        this.setState({ showProvisionLists: false });
    }

    private async setAttendance(key: string) {
        
        this.setState({ buttonDisabled: true });
        const currentUser = await sp.web.currentUser();
        if (!currentUser) {
            console.error(`AttendeeRegister currentUser cannot be found`);
            return;
        }

        const currentListItems = await sp.web.lists.getByTitle(attendanceListTitle).items.filter(`${createdByIdField} eq ${currentUser.Id} and ${titleField} eq '${this.props.calendarIdentifier}'`).getAll();
        if (currentListItems.length > 0) {
            const updateData = {};
            updateData[customAttendanceField] = key;
            await sp.web.lists.getByTitle(attendanceListTitle).items.getById(currentListItems[0].Id).update(updateData);
        }
        else {
            const updateData = {};
            updateData[titleField] = this.props.calendarIdentifier;
            updateData[customAttendanceField] = key;
            await sp.web.lists.getByTitle(attendanceListTitle).items.add(updateData);
        }

        this.resetAttendance();
        this.setState({ buttonDisabled: false });
        await this.loadAttendance();

    }

    public render() {

        // console.log('Rendering');

        const playIcon: IIconProps = this.props.playButtonIcon.length > 0 ? { iconName: this.props.playButtonIcon } : {};
        
        const attendanceRegistered = this.state.attendanceData.filter(attendance => { return attendance[titleField] === this.props.calendarIdentifier; } );
        const attendanceType = this.state.attendanceTypesData.filter(type => { return type[titleField] === (attendanceRegistered.length > 0 && attendanceRegistered[0][customAttendanceField]) || ''; });
        const buttonText = attendanceType.length > 0 ? `Registered: ${attendanceType[0][customAttendanceDisplayNameField]}` : `Register interest`;
        const selectedOption = attendanceType.length > 0 ? `${attendanceType[0][titleField]}` : ``;

        const menuProps: IContextualMenuProps = {
            // For example: disable dismiss if shift key is held down while dismissing
            onDismiss: ev => {
              if (ev && ev.shiftKey) {
                ev.preventDefault();
              }
            },
            onItemClick: (ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): boolean => {
                if (item) {
                    this.setAttendance(item.key);
                    return true;
                }
            },
            items: this.state.attendanceTypesData.map(type => {
                return {
                    key: type[titleField],
                    text: type[customAttendanceDisplayNameField],
                    iconProps: { iconName: type[customIconField] }
                };
            }).concat((attendanceType.length > 0) ? [{
                key: 'remove',
                text: 'Unregister',
                iconProps: { iconName: 'Delete' }
            }] : []),
            directionalHintFixed: true,
        };

        const dragOptions: IDragOptions = {
            moveMenuItemText: 'Move',
            closeMenuItemText: 'Close',
            menu: ContextualMenu
        };

        return (<>

                    <DefaultButton
                        text={buttonText}
                        menuProps={menuProps}
                        // Optional callback to customize menu rendering
                        menuAs={this._getMenu}
                        // Optional callback to do other actions (besides opening the menu) on click
                        // onMenuClick={_onMenuClick}
                        // By default, the ContextualMenu is re-created each time it's shown and destroyed when closed.
                        // Uncomment the next line to hide the ContextualMenu but persist it in the DOM instead.
                        // persistMenu={true}
                        allowDisabledFocus
                        disabled={this.state.buttonDisabled}
                        // checked={checked}
                        />

                {selectedOption === 'ondemand' && this.props.onDemandUrl && (this.props.onDemandUrl.substring(this.props.onDemandUrl.length - 1) !== "=") && 
                    <>
                        <DefaultButton text={this.props.playButtonText} iconProps={playIcon} onClick={() => { window.location.href = this.props.onDemandUrl; }} />
                    </>
                }

                {this.state.includeModal && <Modal
                    isOpen={this.state.showProvisionLists}
                    onDismiss={this.hideProvisionLists}
                    isBlocking={false}
                    // containerClassName={contentStyles.container}
                    dragOptions={dragOptions}
                    >

                    <ProvisionLists _ctx={this.props._ctx} />

                </Modal>}

            </>);
    }

    private _getMenu(props: IContextualMenuProps): JSX.Element {
        // Customize contextual menu with menuAs
        return <ContextualMenu {...props} />;
    }
      
}


