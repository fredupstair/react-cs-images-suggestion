import * as React from 'react';
import { IReactAIsuggestionProps } from './IReactAIsuggestionProps';

import Dropzone from 'react-dropzone'
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { IRectangle } from 'office-ui-fabric-react/lib/Utilities';

import { sp } from "@pnp/sp";
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';

import styles from './ReactAIsuggestion.module.scss';

export interface IReactAIsuggestionState {
  files: any[];
  items: any[];

}
const ROWS_PER_PAGE = 3;
const MAX_ROW_HEIGHT = 250;

export interface IReactAIsuggestionProps {
  context: any
}

export default class ReactAIsuggestion extends React.Component<IReactAIsuggestionProps, IReactAIsuggestionState> {


  private cognitiveServicesKey: string = "";
  private cognitiveServicesVisionUrl: string = "https://westeurope.api.cognitive.microsoft.com/vision/v2.0/tag";


  private _columnCount: number;
  private _columnWidth: number;
  private _rowHeight: number;

  private onDrop(files) {
    this.setState({
      files: files.map(file => Object.assign(file, {
        preview: URL.createObjectURL(file)
      }))
    });

    this._getTagsForImage(files[0])
      .then((tags: any) => {

        const data = [];
        let _data = {};

        // define a search query object matching the SearchQuery interface
        var queryText=tags[0].name;
        sp.search({
          Querytext: queryText
          //RowLimit: 10,
          //EnableInterleaving: true,
          //SelectProperties: ["Title", "Author", "Description", "SiteDescription", "EncodedAbsUrl", "LastModifiedTime", "PictureThumbnailURL", "Path", "Keywords"]
        }).then((r: any) => {//it was SearchResults
          console.log(r.ElapsedTime);
          console.log(r.RowCount);
          console.log(r.PrimarySearchResults);

          r.PrimarySearchResults.forEach(element => {
            console.log(element);
            _data = {
              key: element.UniqueId,
              name: 'Item ' + element.Title,
              tags: element.HitHighlightedSummary,
              thumbnail: element.PictureThumbnailURL
            };

            data.push(_data);
          });

          this.setState({ items: data });

        })
          .catch(error => {
            console.log(error);
          });
      }).catch(error => {

        console.log(error);

      });


  }

  private async _getTagsForImage(fileInfo: any): Promise<string[]> {

    const httpOptions: IHttpClientOptions = this._prepareHttpOptionsForVisionApi(fileInfo);
    const cognitiveResponse: HttpClientResponse = await this.props.context.httpClient.post(this.cognitiveServicesVisionUrl, HttpClient.configurations.v1, httpOptions);
    const cognitiveResponseJSON: any = await cognitiveResponse.json();
    const tags: any = cognitiveResponseJSON.tags;
    return tags;
  }

  private _prepareHttpOptionsForVisionApi(fileInfo: any): IHttpClientOptions {

    const reader = new FileReader();
    reader.onload = (event) => {
      //console.log(reader.result);
    };
    reader.readAsDataURL(fileInfo);
    const httpOptions: IHttpClientOptions = {
      body: fileInfo,
      headers: this._prepareHeadersForVisionApi(),
    };

    return httpOptions;
  }

  private _prepareHeadersForVisionApi(): Headers {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-Type', 'application/octet-stream');
    requestHeaders.append('Cache-Control', 'no-cache');
    requestHeaders.append('Ocp-Apim-Subscription-Key', this.cognitiveServicesKey);

    return requestHeaders;
  }

  componentWillUnmount() {
    // Make sure to revoke the data uris to avoid memory leaks
    this.state.files.forEach(file => URL.revokeObjectURL(file.preview))
  }

  constructor(props: IReactAIsuggestionProps) {
    super(props);
    this.state = {
      files: [],
      items: []
    };
  }

  public render(): React.ReactElement<IReactAIsuggestionProps> {
    const { files } = this.state;

    const thumbs = files.map(file => (
      <div className={styles.thumbDrop} key={file.name}>
        <div className={styles.thumbInnerDrop}>
          <img
            src={file.preview}
            className={styles.imgDrop}
          />
        </div>
      </div>
    ));



    return (
      <div>
        <section>

          <Dropzone
            accept="image/*"
            onDrop={this.onDrop.bind(this)}
          >
            {({ getRootProps, getInputProps }) => (
              <div {...getRootProps()}>
                <input {...getInputProps()} />
                <p>Drop files here</p>
              </div>
            )}
          </Dropzone>
          <aside className={styles.thumbsContainerDrop}>
            {thumbs}
          </aside>
        </section>
        <br /><br /><br /><br /><br /><br />
        <FocusZone>
          <List
            className="ms-ListGridReactAISuggestion"
            items={this.state.items}
            getItemCountForPage={this._getItemCountForPage}
            getPageHeight={this._getPageHeight}
            renderedWindowsAhead={4}
            onRenderCell={this._onRenderCell}
          />
        </FocusZone>

      </div>
    );
  }

  private _getItemCountForPage = (itemIndex: number, surfaceRect: IRectangle): number => {
    if (itemIndex === 0) {
      this._columnCount = Math.ceil(surfaceRect.width / MAX_ROW_HEIGHT);
      this._columnWidth = Math.floor(surfaceRect.width / this._columnCount);
      this._rowHeight = this._columnWidth;
    }

    return this._columnCount * ROWS_PER_PAGE;
  };

  private _getPageHeight = (): number => {
    return this._rowHeight * ROWS_PER_PAGE;
  };

  private _onRenderCell = (item: any, index: number | undefined): JSX.Element => {
    return (
      <div
        className="ms-ListGridReactAISuggestion-tile"
        data-is-focusable={true}
        style={{
          width: 100 / this._columnCount + '%'
        }}
      >
        <div className="ms-ListGridReactAISuggestion-sizer">
          <div className="msListGridReactAISuggestion-padder">
            <img src={item.thumbnail} className="ms-ListGridReactAISuggestion-image" />
            <span className="ms-ListGridReactAISuggestion-label">{item.tags}</span>
          </div>
        </div>
      </div>
    );
  };
}
