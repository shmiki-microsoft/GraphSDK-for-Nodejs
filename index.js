require('dotenv').config();
const msal = require("@azure/msal-node");
const graph = require("@microsoft/microsoft-graph-client");

// MSAL 設定
const msalConfig  = {
    auth: {
      clientId: process.env.CLIENT_ID, 
      authority: process.env.AAD_ENDPOINT + process.env.TENANT_ID,
      clientSecret: process.env.CLIENT_SECRET, 
    },
  };

// ユーザ情報を取得する関数
async function getUserDetails() {
  const cca = new msal.ConfidentialClientApplication(msalConfig );
  
  const clientCredentialRequest = {
    scopes: ["https://graph.microsoft.com/.default"], // API のパーミッション
  };

  try {
    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, response.accessToken);
      },
    });

    // ユーザーの一覧を取得
    const user = await client.api("/users").get();

    // オブジェクトID で指定したユーザーだけ取得
    // const id1 = process.env.USER_ID1;
    // const id2 = process.env.USER_ID2;
    // const user = await client.api("/users").filter(`id in ('${id1}','${id2}')`) .get();
    return user;

  } catch (error) {
    console.log(error);
    throw error;
  }
}

// ユーザ情報を取得する関数
async function getUserDetailsWithSearchQuery() {
  const cca = new msal.ConfidentialClientApplication(msalConfig);
  
  const clientCredentialRequest = {
    scopes: ["https://graph.microsoft.com/.default"], // API のパーミッション
  };

  try {
    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, response.accessToken);
      },
    });

    // ユーザーの一覧を取得
    //https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/HEAD/docs/OtherAPIs.md#query
    const user = await client.api("/users")
    .header('ConsistencyLevel','eventual')
    //.search('displayName:user')
    .search('\"displayName:user\"')
   // .query('$search="displayName:user"')
    .get();
    return user;

  } catch (error) {
    console.log(error);
    throw error;
  }
}

// ユーザー情報を取得して表示
// getUserDetails()
//   .then((users) => {
//     console.log(users);
//   })
//   .catch((error) => {
//     console.log(error);
//   });


// Search Query を使ってユーザー情報を取得し、表示
getUserDetailsWithSearchQuery()
.then((users) => {
  console.log(users);
})
.catch((error) => {
  console.log(error);
});

