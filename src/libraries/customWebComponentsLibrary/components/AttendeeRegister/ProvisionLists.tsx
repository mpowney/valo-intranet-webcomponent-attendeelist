import * as React from 'react';
import { sp } from "@pnp/sp";
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { DetailsList,DetailsListLayoutMode, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields/list";
import "@pnp/sp/views";
import "@pnp/sp/site-users/web";
import styles from './ProvisionLists.module.scss';

export interface IListProvisioned {
    provisioned: boolean;
    name: string;
}

export interface IProvisionListsProps {
    _ctx: any;
}

export interface IProvisionListsState {
    listProvisionedState: IListProvisioned[];
}

const viewName: string = "All Items";
const titleField: string = "Title";
const createdByField: string = "Author";
const attendanceListTitle: string = "SessionAttendance";
const typesListTitle: string = "AttendanceTypes";
const customAttendanceField: string = "AttendanceType";
const customAttendanceDisplayNameField: string = "AttendanceDisplayName";
const customSortOrderField: string = "SortOrder";
const customIconField: string = "Icon";

export class ProvisionLists extends React.Component<IProvisionListsProps, IProvisionListsState> {

    private listsProvisioned: IListProvisioned[] = [
        { provisioned: false, name: attendanceListTitle },
        { provisioned: false, name: typesListTitle }
    ];
    
    constructor(props: IProvisionListsProps) {
        super(props);

        this.clickProvisionLists = this.clickProvisionLists.bind(this);
        this.checkListState = this.checkListState.bind(this);

        sp.setup({spfxContext: props._ctx});

        this.state = {
            listProvisionedState: this.listsProvisioned
        };
      
    }

    public componentDidMount() {
        this.checkListState();
    }
    
    private async clickProvisionLists(ev: React.MouseEvent<HTMLAnchorElement, MouseEvent>) {

        try {
            if (!this.state.listProvisionedState[0].provisioned) await this.provisionAttendanceList();
        } catch {}
        try {
            if (!this.state.listProvisionedState[1].provisioned) await this.provisionTypesList();
        } catch {}

    }

    private async clickProvisionList(listTitle: string) {

        console.log(`Executing clickProvisionList('${listTitle}')`);

        switch (listTitle) {
            case attendanceListTitle: 
                console.log(`Executing provisionAttendanceList()`);
                await this.provisionAttendanceList();
                break;
            case typesListTitle: 
                console.log(`Executing provisionTypesList()`);
                await this.provisionTypesList();
                break;
        }

        await this.checkListState();

    }

    private async checkListState() {
        try {
            const sessionAttendanceProvisioned = await sp.web.lists.getByTitle(attendanceListTitle).get();
            this.listsProvisioned[0].provisioned = true;
            this.setState({ listProvisionedState: this.listsProvisioned });
            console.log(sessionAttendanceProvisioned);
        } catch {
            this.listsProvisioned[0].provisioned = false;
            this.setState({ listProvisionedState: this.listsProvisioned });
        }
        try {
            const attendanceTypeProvisioned = await sp.web.lists.getByTitle(typesListTitle).get();
            this.listsProvisioned[1].provisioned = true;
            this.setState({ listProvisionedState: this.listsProvisioned });
            console.log(attendanceTypeProvisioned);
        } catch {
            this.listsProvisioned[1].provisioned = false;
            this.setState({ listProvisionedState: this.listsProvisioned });
        }

    }

    private async provisionAttendanceList() {

        try {
            await sp.web.lists.add(attendanceListTitle);
            const updateProperties = {
                NoCrawl: true,
                ReadSecurity: 2, // Read items that were created by the user
                WriteSecurity: 2 // Create items and edit items that were created by the user
            };
            await sp.web.lists.getByTitle(attendanceListTitle).update(updateProperties);
            await sp.web.lists.getByTitle(attendanceListTitle).fields.addText(customAttendanceField);
            await sp.web.lists.getByTitle(attendanceListTitle).views.getByTitle(viewName).fields.add(customAttendanceField);
            await sp.web.lists.getByTitle(attendanceListTitle).views.getByTitle(viewName).fields.add(createdByField);

            this.listsProvisioned[0].provisioned = true;
            this.setState({ listProvisionedState: this.listsProvisioned });


        } catch (err) {
            console.log(`Error occurred provisioning ${attendanceListTitle} list: ${JSON.stringify(err)}`);

        }

    }

    private async provisionTypesList() {

        try {
            await sp.web.lists.add(typesListTitle);
            const updateProperties = {
                NoCrawl: true,
            };
            await sp.web.lists.getByTitle(typesListTitle).update(updateProperties);
            await sp.web.lists.getByTitle(typesListTitle).fields.addText(customAttendanceDisplayNameField);
            await sp.web.lists.getByTitle(typesListTitle).fields.addNumber(customSortOrderField);
            await sp.web.lists.getByTitle(typesListTitle).fields.addText(customIconField);
            await sp.web.lists.getByTitle(typesListTitle).views.getByTitle(viewName).fields.add(customAttendanceDisplayNameField);
            await sp.web.lists.getByTitle(typesListTitle).views.getByTitle(viewName).fields.add(customSortOrderField);
            await sp.web.lists.getByTitle(typesListTitle).views.getByTitle(viewName).fields.add(customIconField);

            const listItem1 = {}; listItem1[titleField] = "online"; listItem1[customAttendanceDisplayNameField] = "Attend online"; listItem1[customSortOrderField] = "1"; listItem1[customIconField] = "Video";
            const listItem2 = {}; listItem2[titleField] = "ondemand"; listItem2[customAttendanceDisplayNameField] = "On demand"; listItem2[customSortOrderField] = "2"; listItem2[customIconField] = "ScreenCast";
            await sp.web.lists.getByTitle(typesListTitle).items.add(listItem1);
            await sp.web.lists.getByTitle(typesListTitle).items.add(listItem2);

            // TODO: update default view query to sort by sort order field
            // await sp.web.lists.getByTitle(typesListTitle).views.getByTitle(viewName).query

            this.listsProvisioned[1].provisioned = true;
            this.setState({ listProvisionedState: this.listsProvisioned });

        } catch (err) {
            console.log(`Error occurred provisioning ${typesListTitle} list: ${JSON.stringify(err)}`);

        }

    }

    public render() {

        console.log('Rendering ProvisionLists.tsx');

        const columns: IColumn[] = [
            {
              key: 'provisioned',
              name: 'Provisioned',
              ariaLabel: 'No column operations available for Provisioned column',
              isIconOnly: true,
              fieldName: 'name',
              minWidth: 16,
              maxWidth: 16,
              onRender: (item: IListProvisioned) => {
                return item.provisioned ?
                    <Icon iconName='StatusCircleCheckmark' /> : null;
              },
            },
            {
              key: 'column2',
              name: 'List title',
              fieldName: 'name',
              minWidth: 210,
              maxWidth: 350,
              isRowHeader: true,
              isResizable: true,
              data: 'string',
              isPadded: true,
            },
            {
              key: 'column3',
              name: '',
              fieldName: 'provisioned',
              minWidth: 70,
              maxWidth: 90,
              isResizable: true,
              onRender: (item: IListProvisioned) => {
                return <DefaultButton text="Provision" disabled={item.provisioned} onClick={(ev: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => { this.clickProvisionList(item.name); } } />;
              },
              isPadded: true,
            },
        ];

        return (<div className={styles.provisionListsContainer}>

                    <p>This site isn't set up for registering attendance. Provision the lists below to complete the set up process.</p>

                    <DetailsList
                                items={this.state.listProvisionedState}
                                compact={true}
                                columns={columns}
                                selectionMode={SelectionMode.none}
                                layoutMode={DetailsListLayoutMode.justified}
                                isHeaderVisible={true}
                                selectionPreservedOnEmptyClick={true}
                                enterModalSelectionOnTouch={true}
                                ariaLabelForSelectionColumn="Toggle selection"
                                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                checkButtonAriaLabel="Row checkbox"
                                />


                    <PrimaryButton
                        className={styles.provisionAllButton}
                        text="Provision all"
                        // Optional callback to customize menu rendering
                        // Optional callback to do other actions (besides opening the menu) on click
                        // onMenuClick={_onMenuClick}
                        // By default, the ContextualMenu is re-created each time it's shown and destroyed when closed.
                        // Uncomment the next line to hide the ContextualMenu but persist it in the DOM instead.
                        // persistMenu={true}
                        allowDisabledFocus
                        // disabled={disabled}
                        // checked={checked}
                        onClick={this.clickProvisionLists}
                        />

            </div>);
    }

}


