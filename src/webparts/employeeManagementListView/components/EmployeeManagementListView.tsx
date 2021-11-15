import * as React from "react";
import styles from "./EmployeeManagementListView.module.scss";
import { IEmployeeManagementListViewProps } from "./IEmployeeManagementListViewProps";
import { escape } from "@microsoft/sp-lodash-subset";

// REQUIRED IMPORTS
///////////////////////////////////////////////////////////////////////////////
import {
  Announced,
  DetailsList,
  IColumn,
  MarqueeSelection,
  TextField,
  Selection,
  DetailsListLayoutMode,
  ITextFieldStyles,
  mergeStyles,
} from "office-ui-fabric-react";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IDetailsListBasicExampleItem } from "./IDetailsListBasicExampleItem";
///////////////////////////////////////////////////////////////////////////////

const exampleChildClass = mergeStyles({
  display: "block",
  marginBottom: "10px",
});

const textFieldStyles: Partial<ITextFieldStyles> = {
  root: { maxWidth: "300px" },
};

export default class EmployeeManagementListView extends React.Component<
  IEmployeeManagementListViewProps,
  { items: IDetailsListBasicExampleItem[]; selectionDetails: string }
> {
  //////////// default for page link
  // required in development
  private w = Web(this.props.webUrl + "/sites/Maitri");

  // required in production
  // w = Web(this.props.webUrl);

  private url = location.search;
  private params = new URLSearchParams(this.url);
  private id = this.params.get("spid");

  ////////////

  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];

  constructor(props: IEmployeeManagementListViewProps, state: any) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    // get all list items and log in console.
    let fitems: any[];
    this._allItems = [];

    // Populate with items for demos.
    // for (let i = 0; i < 200; i++) {
    //   this._allItems.push({
    //     ID: i,
    //     EmpName: "Item " + i,
    //     JobDescription: i.toString(),
    //     Gender: "",
    //   });
    // }

    this._columns = [
      {
        key: "column1",
        name: "Employee Name",
        fieldName: "EmpName",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column2",
        name: "Job Description",
        fieldName: "JobDescription",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column3",
        name: "Gender",
        fieldName: "Gender",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column4",
        name: "Date Of Birth",
        fieldName: "DOB",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "column5",
        name: "Address",
        fieldName: "Address",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
    ];

    this.state = {
      items: this._allItems,
      selectionDetails: this._getSelectionDetails(),
    };
  }

  public render(): React.ReactElement<IEmployeeManagementListViewProps> {
    return (
      <div>
        <div className={exampleChildClass}>{this.state.selectionDetails}</div>
        <Announced message={this.state.selectionDetails} />
        <TextField
          className={exampleChildClass}
          label="Filter by name:"
          onChange={this._onFilter}
          styles={textFieldStyles}
        />
        <Announced
          message={`Number of items after filter applied: ${this.state.items.length}.`}
        />
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={this.state.items}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="select row"
            onItemInvoked={this._onItemInvoked}
          />
        </MarqueeSelection>
      </div>
    );
  }

  public componentDidMount = () => {
    this._getAllEmps();
  };

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return (
          "1 item selected: " +
          (this._selection.getSelection()[0] as IDetailsListBasicExampleItem)
            .EmpName
        );
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onFilter = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text: string
  ): void => {
    this.setState({
      items: text
        ? this._allItems.filter(
            (i) => i.EmpName.toLowerCase().indexOf(text) > -1
          )
        : this._allItems,
    });
  };

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    alert(
      `Item invoked: ${item.EmpName}, with id: ${item.ID}, form opening in new tab`
    );
    window.open(
      `https://therisav.sharepoint.com/SitePages/CURD-Form.aspx?spid=${item.ID}`,
      "CURD FORM"
    );
  };

  private _getAllEmps = () => {
    /* Fetches all emps from sp Employee_Master table 
    1. setting the detailist items
    */

    // basic usage
    this.w.lists
      .getByTitle("Employee_Master")
      .items.getAll()
      .then((allItems: any[]) => {
        // // how many did we get
        // console.log(
        //   allItems.length,
        //   ' emp records fetched from emp master succesfully....'
        // );

        // console.log(allItems);

        // setting detaillist item list
        allItems.map((i) => {
          this.setState({
            items: [
              ...this.state.items,
              {
                key: i.Id,
                ID: i.Title,
                EmpName: i.Employee_Name,
                JobDescription: "",
                Gender: i.Gender,
                DOB: i.Date_of_Birth,
                Address: i.Address,
              },
            ],
          });
        });
      });
  };
}
