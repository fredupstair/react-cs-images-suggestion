import * as React from 'react';
import { Fragment } from 'react';
import { IReactCSsuggestionProps } from './IReactCSsuggestionProps';

import Dropzone from 'react-dropzone'
import { FocusZone, TextField, List, IRectangle } from 'office-ui-fabric-react';

import { sp, SearchQueryBuilder, SearchQuery } from "@pnp/sp";
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
import { String } from 'typescript-string-operations';
import * as strings from 'ReactCSsuggestionWebPartStrings';

import styles from './ReactCSsuggestion.module.scss';
import { stringIsNullOrEmpty } from '@pnp/common';

export interface IReactCSsuggestionState {
  files: any[];
  items: any[];
  tagsResult: string;
}
const ROWS_PER_PAGE = 3;
const MAX_ROW_HEIGHT = 250;

export interface IReactCSsuggestionProps {
  context: any
}

export default class ReactCSsuggestion extends React.Component<IReactCSsuggestionProps, IReactCSsuggestionState> {


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
      .then((tags: string[]) => {

        //first 3 relevants tag, join by coma and build  FQL query
        //https://docs.microsoft.com/en-us/sharepoint/dev/general-development/fast-query-language-fql-syntax-reference
        var tagsText = "";
        var tagsQuery = "";
        tags.slice(0, 3).forEach(element => {
          tagsText += String.Format("{0}; ", element['name']);
          tagsQuery += String.Format('"{0}",', element['name']);
        });

        tagsQuery = tagsQuery.substr(0, tagsQuery.length - 1);
        this.setState({ tagsResult: tagsText });

        //var q = 'and(near(' + tagsQuery + 'N=2),filetype:"jpg")';
        var q = String.Format('and(or({0}),filetype:"jpg")', tagsQuery);
        //var q = tagsQuery;
        const _searchQuerySettings: SearchQuery = {
          TrimDuplicates: true,
          EnableFQL: true,
          RowLimit: 50,
          SelectProperties: ["Title", "SPWebUrl", "DefaultEncodingURL", "HitHighlightedSummary", "KeywordsOWSMTXT"]
        }

        let query = SearchQueryBuilder(q, _searchQuerySettings);


        sp.search(query).then((r: any) => {//it was SearchResults
          console.log(r.ElapsedTime);
          console.log(r.RowCount);
          console.log(r.PrimarySearchResults);
          const data = [];
          let _data = {};
          r.PrimarySearchResults.forEach(element => {
            console.log(element);
            _data = {
              key: element.UniqueId,
              name: 'Item ' + element.Title,
              tags: element.KeywordsOWSMTXT,
              thumbnail: element.DefaultEncodingURL
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

  constructor(props: IReactCSsuggestionProps) {
    super(props);
    this.state = {
      files: [],
      items: [],
      tagsResult: ""
    };
  }

  public render(): React.ReactElement<IReactCSsuggestionProps> {
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
                <p className={styles.pDrop}>{strings.DropFileHere}</p>
              </div>
            )}
          </Dropzone>
          <TextField label={strings.TagsResult} readOnly={true} value={this.state.tagsResult} />
          <aside className={styles.thumbsContainerDrop}>
            {thumbs}
          </aside>
        </section>
        <br />
        <FocusZone>
          <List
            className="ms-ListGridReactCSsuggestion"
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
        className="ms-ListGridReactCSsuggestion-tile"
        data-is-focusable={true}
        style={{
          width: 100 / this._columnCount + '%'
        }}
      >
        <div className="ms-ListGridReactCSsuggestion-sizer">
          <div className="msListGridReactCSsuggestion-padder">
            <img src={item.thumbnail} className="ms-ListGridReactCSsuggestion-image" />
            <span className="ms-ListGridReactCSsuggestion-label"><Fragment>{item.tags}</Fragment></span>
          </div>
        </div>
      </div>
    );
  };
}
