import * as React from 'react';
import styles from './Search.module.scss';
import { ISearchProps } from './ISearchProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import pnp, { Web, sp } from 'sp-pnp-js';

export default class Search extends React.Component<ISearchProps, {}> {
  public state = {
    rawitems: [],
    columns: [],
    refitems:[]
  };
  public render(): React.ReactElement<ISearchProps> {
    return (
      <div className={styles.search}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
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
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
  public componentDidMount() {
    // this.getData();
    this.setState({ columns: this.columsCreate(['Id', 'Title', 'Color']) });
    this.searchList();

  }
  public componentWillReceiveProps(){
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
    //  https://cupcuper.sharepoint.com/search/_api/search/query?querytext='colors'&clienttype='ContentSearchRegular'
    console.log(this.props.listname);
    fetch(`https://cupcuper.sharepoint.com/search/_api/search/query?querytext='b40ba88e-c159-47ad-bef0-43ea539cf24c'&selectproperties='title%2cRefinableString50'&clienttype='ContentSearchRegular'`, {
      method: 'get',
      headers: {
        'accept': "application/json;odata=verbose",
        'content-type': "application/json;odata=verbose", 
      }
    }).then((response)=>response.json()).then((d)=>{
      console.log(d.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results);
      return d.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
    }
    ).then((data)=> {
      console.log(data);
      data.forEach(item => {
        // console.log(item.Cells.results[2].Value)
        refitems.push({
          Id: item.Cells.results[1].Value,
          Title: item.Cells.results[2].Value,
          Color: item.Cells.results[3].Value
        }),this.setState({refitems: refitems});
        
      });
    }).then(()=>console.log(refitems));
  }
  private getData(): void {
    // let uri = this.props.siteurl || 'sites/dev1';
    // console.log(this.props.siteurl);
    let guid = /*this.props.listdropdown ||*/ '78d16b51-cb43-4802-a42b-3713059cd18b';
    let columns = this.columsCreate(['Id', 'Title']);
    this.setState({ columns: columns });
    console.log(this.props.context.pageContext.web);
    let wep = new Web(this.props.context.pageContext.web.absoluteUrl + '/sites/dev1');
    // pnp.sp.web.lists.filter('Hidden eq false').get().then((li) => console.log(li));
    wep.lists.getById(guid).items.get().then
      ((response) => {
        this.setState({ rawitems: response });
        console.log('writing');
      }
      ).then(() => console.log(this.state));
  }
}
