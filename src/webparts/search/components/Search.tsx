import * as React from 'react';
import styles from './Search.module.scss';
import { ISearchProps } from './ISearchProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import pnp, { Web, sp } from 'sp-pnp-js';

export default class Search extends React.Component<ISearchProps, {}> {
  public state = {
    columns: [],
    refitems: []
  };
  public render(): React.ReactElement<ISearchProps> {
    return (
      <div className={styles.search}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <DetailsList
                items={this.state.refitems}
                compact={false}
                columns={this.state.columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selectionPreservedOnEmptyClick={true}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              />
              <p className={styles.description}>{escape(this.props.listname)}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
  public componentDidMount() {
    this.setState({ columns: this.columsCreate(['Id', 'Title', 'Color']) });
    this.searchList();

  }
  private columsCreate(arraySelect: Array<any>): Array<IColumn> {
    const columns: IColumn[] = [];
    arraySelect.forEach((el, index) => {
      columns.push({
        key: `column${index}`,
        name: el,
        fieldName: el,
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
      });
    });
    return columns;
  }
  private searchList() {
    let refitems = [];
    fetch(`https://cupcuper.sharepoint.com/search/_api/search/query?querytext='b40ba88e-c159-47ad-bef0-43ea539cf24c'&selectproperties='title%2cRefinableString50'&clienttype='ContentSearchRegular'`, {
      method: 'get',
      headers: {
        'accept': "application/json;odata=verbose",
        'content-type': "application/json;odata=verbose",
      }
    }).then((response) => response.json()).then((d) => {
      return d.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
    }
    ).then((data) => {
      console.log(data);
      data.forEach(item => {
        refitems.push({
          Id: item.Cells.results[1].Value,
          Title: item.Cells.results[2].Value,
          Color: item.Cells.results[3].Value
        }), this.setState({ refitems: refitems });

      });
    });
  }
}
