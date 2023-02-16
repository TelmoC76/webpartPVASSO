import * as React from "react";
import jwt_decode from "jwt-decode";
import * as ReactWebChat from 'botframework-webchat';
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  AuthorizationUrlRequest,
  AccountInfo
} from "@azure/msal-browser";
//import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';



export interface IChatbotProps {
  tenantId: string,
  redirectUri: string,
  botid: string,
  userName: string,
  tokenOpenId: string,
  clientId: string,
  botScope: string
  userId: string
}

export interface IChatbotState {
  tokenOpenId: string;
}




export const PVAChatbotDialog: React.FunctionComponent<IChatbotProps> = (props: IChatbotProps) => {

  var conversationId = '';




  const theURL = "https://powerva.microsoft.com/api/botmanagement/v1/directline/directlinetoken?botId=" + props.botid;

  function arrayBufferToBase64(buffer: ArrayBuffer) {
    alert(buffer.byteLength);
    var binary = '';
    var bytes = new Uint8Array(buffer);
    var len = bytes.byteLength;
    alert(len);
    for (var i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  function uploadFile(_input: any) {
    fetch(_input.url).then(response => response.arrayBuffer()).then(buffer => {
      fetch('https://prod-92.westus.logic.azure.com:443/workflows/6da685f3348a4fa093164f24f671e82f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=AWgXlfQJ_vE3gTK_WStzQ6xqZAmCZKSKK4TppEu_4ho',
        {
          method: 'POST',
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
          },
          body: "{ \"conversationid\":\"" + conversationId + "\",\"base64Content\":\"" + arrayBufferToBase64(buffer) + "\"}"
        }).then(response => response.json()).then(response => console.log(JSON.stringify(response)))

    })

  }







  fetch(theURL)
    .then(response => response.json())
    .then(conversationInfo => {
      document.getElementById("loading-spinner").style.display = 'none';
      document.getElementById("webchat").style.minHeight = '50vh';
      const directLine = ReactWebChat.createDirectLine({ token: conversationInfo.token, });
      const store = ReactWebChat.createStore(
        {},
        ({ dispatch }: { dispatch: any }) => (next: any) => (action: any) => {
          const { type } = action;
          if (action.type === 'DIRECT_LINE/CONNECT_FULFILLED') {
            dispatch({
              type: 'WEB_CHAT/SEND_EVENT',
              payload: {
                name: 'startConversation',
                type: 'event',
                value: { text: "hello" }
              }
            });
          }
          else if (type === 'WEB_CHAT/SEND_FILES') {
            (async function () {
              // Tells the bot that you are uploading files. This is optional.
              dispatch({ type: 'WEB_CHAT/SEND_TYPING' });

              await Promise.all(action.payload.files.map(({ name, url }: { name: string, url: string }) => uploadFile({ name, url })));

              // In order for Power VA to process a file, we need to pass the blob URL as a String within a message.
              // We intercept the Event and rather than sending it as an attachment we are converting the URL to a string
              // and dispatching a message.
              dispatch({
                type: 'WEB_CHAT/SEND_MESSAGE_BACK',
                payload: {
                  text: "Your files were successfully uploaded"
                }
              });

            })().catch(err => console.error(err));
          }
          else if (action.type === 'DIRECT_LINE/INCOMING_ACTIVITY') {
            const activity = action.payload.activity;


            if (activity &&
              activity.attachments &&
              activity.attachments[0] &&
              activity.attachments[0].contentType === 'application/vnd.microsoft.card.oauth' &&
              activity.attachments[0].content.tokenExchangeResource) {
              if (props.tokenOpenId && props.tokenOpenId != "") {
                const tokenx = props.tokenOpenId;
                console.log(activity.attachments[0].content.tokenExchangeResource.id);
                console.log(activity.attachments[0].content.tokenExchangeResource.uri);
                var decoded = jwt_decode(tokenx);
                console.log(decoded);

                directLine.postActivity({
                  //@ts-ignore: Activity Schema is valid for DirectLine, but not the botframework-webchat package
                  type: "invoke",
                  name: 'signin/tokenExchange',
                  value: {
                    id: activity.attachments[0].content.tokenExchangeResource.id,
                    connectionName: activity.attachments[0].content.connectionName,
                    tokenx
                  },
                  "from": {
                    id: props.userName,
                    name: props.userId,
                    role: "user"
                  }
                }).subscribe(
                  id => {
                    if (id === 'retry') {
                      // Bot was not able to handle the invoke, so display the oauthCard
                      return next(action);
                    }
                    // Tokenexchange successful and we do not display the oauthCard
                  },
                  error => {
                    // An error occurred to display the oauthCard
                    return next(action);
                  }
                );
              }
              else {
                return next(action);
              }
            }
          }

          else {
            return next(action);
          }


        });

      const styleOptions = {
        // Add styleOptions to customize Web Chat canvas
        //hideUploadButton: true,
        botAvatarImage: 'https://bot-framework.azureedge.net/bot-icons-v1/6ab9b101-b65c-4357-9e9f-915cbf313a14_2K5Bt02aW8egEb97fxAgh7vqChK4UV3Nh3Lw3YYArhEKR8mB.png',
        botAvatarInitials: 'Bot',
        userAvatarImage: 'https://content.powerapps.com/resource/makerx/static/media/user.0d06c38a.svg',
        userAvatarInitials: 'User'
      };
      ReactWebChat.renderWebChat(
        {
          directLine: directLine,
          store: store,
          userID: props.userId,
          styleOptions: styleOptions
        },
        document.getElementById('webchat')
      );
      conversationId = conversationInfo.conversationId;
      console.log(conversationId);
    })
    .catch(err => console.error("An error occurred: " + err));

  return (
    <>

      <div id="chatContainer" style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
        <div id="webchat" role="main" style={{ width: "100%", height: "0rem" }}></div>
        <Spinner id="loading-spinner" label="Loading..." style={{ paddingTop: "1rem", paddingBottom: "1rem" }} />
      </div>

    </>
  );
};

export default class Chatbot extends React.Component<IChatbotProps, IChatbotState> {

  private myMSALObj: PublicClientApplication;



  private ssoRequest = {
    scopes: ["User.Read", "openid", "profile"],
    loginHint: ""
  };


  constructor(props: IChatbotProps) {
    super(props);
    this.ssoRequest = {
      scopes: ["User.Read", "openid", "profile"],
      loginHint: props.userName
    }
    this.state = { tokenOpenId: "" }

  }


  public componentDidMount(): void {
    const msalConfig = {
      auth: {
        authority: "https://login.microsoftonline.com/" + this.props.tenantId,
        clientId: this.props.clientId,
        redirectUri: this.props.redirectUri
      }
    };

    this.myMSALObj = new PublicClientApplication(msalConfig);
    this.myMSALObj.handleRedirectPromise().then((tokenResponse) => {
      if (tokenResponse !== null) {
        const naccess_token = tokenResponse.accessToken;
        console.log(naccess_token);
      } else {
        // In case we would like to directly load data in case of NO redirect:
        const currentAccounts = this.myMSALObj.getAllAccounts();

        if (currentAccounts !== null && currentAccounts.length > 0) {
          this.handleLoggedInUser(currentAccounts);
          this.exchangeToken(currentAccounts);
        }
        else {
          this.loginForAccessTokenByMSAL()
            .then((token) => {
              console.log(token);
            });
        }
      }
    }).catch((error) => {
      console.log(error);
      return null;
    });
  }





  public render(): JSX.Element {
    return (
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", paddingBottom: "1rem" }}>
        <PVAChatbotDialog botScope={this.props.botScope} clientId={this.props.clientId} tokenOpenId={this.state.tokenOpenId} tenantId={this.props.tenantId} userName={this.props.userName} userId={this.props.userId} redirectUri={this.props.redirectUri} botid={this.props.botid} />
      </div>
    );
  }

  private async loginForAccessTokenByMSAL(): Promise<string> {


    return this.myMSALObj.ssoSilent(this.ssoRequest).then((response) => {
      return response.accessToken;
    }).catch((silentError) => {
      console.log(silentError);
      if (silentError instanceof InteractionRequiredAuthError) {
        return this.myMSALObj.loginPopup(this.ssoRequest)
          .then((response) => {
            return response.accessToken;
          })
          .catch(popupError => {
            if (popupError.message.indexOf('popup_window_error') > -1) { // Popups are blocked
              return this.redirectLogin(null);
            }
          });
      } else {
        return null;
      }
    });
  }


  private handleLoggedInUser(currentAccounts: AccountInfo[]) {


    let accountObj = null;
    if (currentAccounts === null) {
      // No user signed in
      return;
    } else if (currentAccounts.length > 1) {
      // More than one user is authenticated, get current one 
      accountObj = this.myMSALObj.getAccountByUsername(this.props.userName);
    } else {
      accountObj = currentAccounts[0];
    }
    if (accountObj !== null) {
      this.acquireAccessToken(null, accountObj)
        .then((accessToken) => {
          var decoded = jwt_decode(accessToken);
          console.log(decoded);
        });
    }
  }

  private async acquireAccessToken(ssoRequest: AuthorizationUrlRequest, account: AccountInfo): Promise<string> {
    const accessTokenRequest = {
      scopes: this.ssoRequest.scopes,
      account: account
    };
    return this.myMSALObj.acquireTokenSilent(accessTokenRequest).then((val) => {
      return val.accessToken;
    }).catch((errorinternal) => {
      console.log(errorinternal);
      return null;
    });
  }

  private redirectLogin(ssoRequest: AuthorizationUrlRequest): Promise<string> {
    try {

      this.myMSALObj.loginRedirect(this.ssoRequest)
        .then(() => {
          return Promise.resolve('');
        });

    } catch (err) {
      console.log(err);
      return null;
    }
  }

  private exchangeToken(accounts: AccountInfo[]) {
    let user = accounts;

    if (user.length <= 0) {
      return null
    }

    const tokenRequest = {
      scopes: [this.props.botScope],
      account: accounts[0]
    };
    console.log("bot scope used in token - " + this.props.botScope);

    return this.myMSALObj.acquireTokenSilent(tokenRequest).then((val) => {
      this.setState({ tokenOpenId: val.accessToken })
      return val.accessToken;
    }).catch((errorinternal) => {
      console.log(errorinternal);
      return null;
    });

  }
}  