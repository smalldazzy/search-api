import * as React from 'react';
import styles from './Search.module.scss';
import { ISearchProps } from './ISearchProps';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { taxonomy, ITermStore, ITermSet } from "@pnp/sp-taxonomy";
import RefButtons from './RefButtons';


export default class Search extends React.Component<ISearchProps, {}> {
  public state = {
    columns: [],
    refitems: [],
    terms: []
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
              <RefButtons
                terms={this.state.terms}
                search={this.searchList}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
  public componentDidMount() {
    this.searchList = this.searchList.bind(this);
    this.setState({ columns: this.columsCreate(['Id', 'Title', 'Color']) });
    this.searchList('');
    this.getTax();
  }
  private async getTax() {
    const store: ITermStore = taxonomy.termStores.getByName("Taxonomy_MbJwt7/iPAiwadr9aBselw==");
    const set: ITermSet = store.getTermSetById("b40ba88e-c159-47ad-bef0-43ea539cf24c");
    set.terms.get().then((response) => {
      let colors = [];
      response.forEach((item) => colors.push(item.Name));
      this.setState({ terms: colors });
    }
    ).then(() => console.log(this.state));
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
  private searchList(refinerquery: string) {
    let refitems = [];
    console.log(refinerquery);
    let url = `https://cupcuper.sharepoint.com/search/_api/search/query?querytext='b40ba88e-c159-47ad-bef0-43ea539cf24c'&selectproperties='title%2cRefinableString50'&clienttype='ContentSearchRegular'`;
    if (refinerquery.length > 0) {
      url = `https://cupcuper.sharepoint.com/search/_api/search/query?querytext='b40ba88e-c159-47ad-bef0-43ea539cf24c'&selectproperties='title%2cRefinableString50'&refinementfilters='RefinableString50:equals(${refinerquery})'&clienttype='ContentSearchRegular'`;
    }
    fetch(url, {
      method: 'get',
      headers: {
        'accept': "application/json;odata=verbose",
        'content-type': "application/json;odata=verbose",
      }
    }).then((response) => response.json()).then((d) => {
      return d.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
    }
    ).then((data) => {
      if (data.length>0) {
        data.forEach(item => {
          refitems.push({
            Id: item.Cells.results[1].Value,
            Title: item.Cells.results[2].Value,
            Color: item.Cells.results[3].Value
          }), this.setState({ refitems: refitems });
        });
      } else {
        alert('No elements found');
      }
    });
  }
}
