import {Version} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';
//import {escape, set} from '@microsoft/sp-lodash-subset';
import * as strings from 'ExternalSiteShareCleanerWebPartStrings';
export interface IExternalSiteShareCleanerWebPartProps {
  Office365Launcher: boolean;
  ShareSiteButton: boolean;
  SendByEmailButton: boolean;
  FooterBar: boolean;
  extraQuerySelectorValues: string;
  description: string;
  showDebug: boolean;
}

export default class ExternalSiteShareCleanerWebPart extends BaseClientSideWebPart<IExternalSiteShareCleanerWebPartProps> {
  private elementsToHide = [];
  public errorLog;
  private maxHideAttempts = 20;

  public render(): void {
    this.configureElements(()=>{
      if (document.location.href.indexOf("Mode=Edit") === -1) {
        this.errorLog = '';
        console.log('ExternalSiteShareCleaner is Activating');
        this.activate();
      } else {
        this.domElement.innerHTML = `<div><h1>External Site Share Cleaner</h1><p>This webpart is hidden when not in edit mode<br>If you are seeing this on a page without edit mode, refresh the page.</p></div>`;
      }
    });

  }



  protected activate() {
    if (this.maxHideAttempts > 0) {
      this.maxHideAttempts--;
      setTimeout(() => {
        if (this.performOperations()) {
          console.log('ExternalSiteShareCleaner is now Active');
          //console.log(this.elementsToHide);
        } else {
          this.activate();
        }
      }, 500);
    } else {
      for (let elementToHide of this.elementsToHide){
        if (elementToHide.hidden === false){
          this.errorLog+=elementToHide.lastError + '\n';
        }
      }
      console.log('ExternalSiteShareCleaner Failed after 20 attempts to clean up the page');
      if (this.properties.showDebug) {
        console.log(this.errorLog);
        this.domElement.innerHTML = '<h3>ExternalSiteShareCleaner Error Log</h3><textarea style="width:100%; height:100px;">' + this.errorLog + '</textarea>This can be disabled in the configuration pane.';
      }
    }
  }

  protected performOperations() {
    let allOpsComplete = true;
    for (let elementToHide of this.elementsToHide){
      if (!elementToHide.hidden){
        let result = this.hideByQuerySelector(elementToHide.querySelector);
        if (!result.hidden){
          elementToHide.lastError = result.lastError;
          allOpsComplete = false;
        } else {
          elementToHide.hidden = true;
        }

      }
    }
    return allOpsComplete;
  }

  protected hideByQuerySelector(selector) {
    //let element = <HTMLElement>document.querySelector(selector);
    let result = {hidden: false, lastError: ''};
    try {
      let elements = document.querySelectorAll(selector);
      if (elements.length < 1) {
        result.lastError = 'Could not find any element with the selector "' + selector + '"';
        return result;
      } else {
        for (let i = 0; i < elements.length; i++) {
          let element = <HTMLElement>elements[i];
            //console.log(element);
            element.style.display = "none";
            result.hidden = true;
        }
        return result;
      }
    } catch (e) {
      result.lastError = e;
      if (this.maxHideAttempts < 1) console.log(e);
      return result;
    }
  }

  protected configureElements(next){
    //Grab the default Elements from the json file
    let elementsFromJSON = <Array<Object>>require('./defaultElements.json');
    for (let eIdx = 0; eIdx < elementsFromJSON.length; eIdx++){
      let e = <any> elementsFromJSON[eIdx];
      if (this.properties[e.propertyName]){
        this.elementsToHide.push(e);
      }
    }
    //add extras from extra prop
    let extras = this.properties.extraQuerySelectorValues.split(',');
    for (let i = 0; i < extras.length; i++){
      let customObj = {
        "name":"Custom " + i,
        "propertyName": "",
        "querySelector": extras[i],
        "hidden": false
      };
      this.elementsToHide.push(customObj);
    }
    next();
  }



  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneCheckbox('Office365Launcher', {
                  text: strings.Office365LauncherBarLabel,
                  checked: false,
                  disabled: false
                }),
                PropertyPaneCheckbox('ShareSiteButton', {
                  text: strings.ShareSiteButtonLabel,
                  checked: false,
                  disabled: false
                }),
                PropertyPaneCheckbox('SendByEmailButton', {
                  text: strings.SendByEmailButtonLabel,
                  checked: false,
                  disabled: false
                }),
                PropertyPaneCheckbox('FooterBar', {
                  text: strings.FooterBarLabel,
                  checked: false,
                  disabled: false
                })
              ]
            },
            {
              groupName: strings.CustomGroupName,
              groupFields: [
                PropertyPaneTextField('extraQuerySelectorValues', {
                  multiline: true,
                  label: strings.extraQuerySelectorValuesLabel
                }),
                PropertyPaneCheckbox('showDebug', {
                  text: strings.showDebugLabel,
                  checked: false,
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
