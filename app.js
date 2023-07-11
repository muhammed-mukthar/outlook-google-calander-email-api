const express = require("express");
const { google } = require("googleapis");
const app = express();
const port = 4000; // Change to the desired port
const cors = require("cors");
const { PublicClientApplication, InteractionRequiredAuthError } = require("@azure/msal-node");
const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");

// Configuration
const credentials = require("./credentials.json");
const gmailScopes = [
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/calendar.readonly",
];

const outlookScopes = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Calendars.Read",
];

const mongoose = require("mongoose");

// MongoDB configuration
const MONGODB_URI =
  "mongodb+srv://mukthar:mukthar@cluster0.zd6myfd.mongodb.net/tokenStore?retryWrites=true&w=majority"; //

// Create a schema for storing access tokens
const tokenSchema = new mongoose.Schema({
  provider: String,
  accessToken: String,
});

// Create a model for the access tokens
const Token = mongoose.model("Token", tokenSchema);

// Connect to MongoDB
mongoose
  .connect(MONGODB_URI, { useNewUrlParser: true, useUnifiedTopology: true })
  .then(() => console.log("Connected to MongoDB"))
  .catch((error) => console.error("Error connecting to MongoDB:", error));
// Gmail configuration
const {
  client_secret: gmailClientSecret,
  client_id: gmailClientId,
  redirect_uris: gmailRedirectUris,
} = credentials.installed;
const gmailOAuth2Client = new google.auth.OAuth2(
  gmailClientId,
  gmailClientSecret,
  gmailRedirectUris[0]
);

const calendarOAuth2Client = new google.auth.OAuth2(
  gmailClientId,
  gmailClientSecret,
  gmailRedirectUris[0]
);
// Outlook configuration
const {
  client_secret: outlookClientSecret,
  client_id: outlookClientId,
  redirect_uris: outlookRedirectUris,
  tenant_id: tenantId,
} = credentials.outlook;
const outlookClientCredential = new ClientSecretCredential(
  tenantId,
  outlookClientId,
  outlookClientSecret
);

console.log(outlookClientSecret, "dssd");

const msalConfig = {
  auth: {
    clientId: outlookClientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
  },
};
const msalClient = new PublicClientApplication(msalConfig);

app.use(cors());

// Gmail Routes

// Get the Gmail authorization URL
const gmailAuthUrl = gmailOAuth2Client.generateAuthUrl({
  access_type: "offline",
  scope: gmailScopes,
});

// Set up a route to initiate the Gmail authentication flow
app.get("/gmail/auth", (req, res) => {
  res.redirect(gmailAuthUrl);
});

// Set up a route to handle the Gmail authorization callback
app.get("/gmail/auth/callback", async (req, res) => {
  const code = req.query.code;

  try {
    // Exchange the Gmail authorization code for an access token
    const { tokens } = await gmailOAuth2Client.getToken(code);
    gmailOAuth2Client.setCredentials(tokens);

    // Save the Gmail access token to the database
    const gmailToken = new Token({
      provider: "gmail",
      accessToken: tokens.access_token,
    });
    await gmailToken.save();
    res.json({ accessToken: tokens.access_token });
  } catch (error) {
    console.error("Error exchanging Gmail authorization code:", error);
    res.status(500).json({ error: "Failed to exchange authorization code for access token" });
  }
});

// Set up a route to logout from Gmail
app.get("/gmail/logout", async (req, res) => {
  try {
    // Clear the Gmail access token from the database
    await Token.deleteOne({ provider: "gmail" });

    // Revoke the Gmail access token
    await gmailOAuth2Client.revokeCredentials();

    console.log("Gmail access token revoked.");
    res.send("Logged out from Gmail.");
  } catch (error) {
    console.error("Error revoking Gmail access token:", error);
    res.status(500).send("An error occurred during Gmail logout.");
  }
});

// Set up a route to get the list of Gmail emails
app.get("/gmail/emails", async (req, res) => {
  try {
    let gmailAccessToken = await Token.find({ provider: "gmail" });
    // Check if the user is authenticated with Gmail
    if (!gmailAccessToken[0]) {
      return res.status(401).send("User not authenticated with Gmail.");
    }

    // Create the Gmail API client with the access token
    const gmail = google.gmail({
      version: "v1",
      auth: gmailOAuth2Client,
    });

    // Get the list of Gmail emails
    const response = await gmail.users.messages.list({
      userId: "me",
    });

    const emails = response.data.messages;

    // Iterate over each email and retrieve additional details
    const emailDetails = await Promise.all(
      emails.map(async (email) => {
        const emailId = email.id;
        const emailData = await gmail.users.messages.get({
          userId: "me",
          id: emailId,
          format: "full",
          metadataHeaders: ["Subject", "From"],
        });

        const subject = emailData.data.payload.headers.find(
          (header) => header.name === "Subject"
        ).value;
        const sender = emailData.data.payload.headers.find(
          (header) => header.name === "From"
        ).value;

        return {
          subject,
          sender,
          emailId,
        };
      })
    );

    res.json(emailDetails);
  } catch (error) {
    console.error("Error retrieving Gmail emails:", error);
    res.status(500).send("An error occurred while retrieving Gmail emails.");
  }
});

// Set up a route to get a specific Gmail email by ID
app.get("/gmail/emails/:id", async (req, res) => {
  try {
    // Check if the user is authenticated with Gmail
    let gmailAccessToken = await Token.find({ provider: "gmail" });
    // Check if the user is authenticated with Gmail
    if (!gmailAccessToken[0]) {
      return res.status(401).send("User not authenticated with Gmail.");
    }

    const emailId = req.params.id; // Get the email ID from the URL

    // Create the Gmail API client with the access token
    const gmail = google.gmail({
      version: "v1",
      auth: gmailOAuth2Client,
    });

    // Get the specific Gmail email by ID
    const response = await gmail.users.messages.get({
      userId: "me",
      id: emailId,
    });

    const email = response.data;

    res.json(email);
  } catch (error) {
    console.error("Error retrieving Gmail email:", error);
    res.status(500).send("An error occurred while retrieving the Gmail email.");
  }
});

// Outlook Routes

// Set up a route to initiate the Outlook authentication flow
app.get("/outlook/auth", async (req, res) => {
  try {
    const outlookAuthUrl = await msalClient.getAuthCodeUrl({
      scopes: outlookScopes,
      redirectUri: outlookRedirectUris[0],
    });

    res.redirect(outlookAuthUrl);
  } catch (error) {
    console.error("Error generating Outlook authorization URL:", error);
    res.status(500).send("An error occurred during Outlook authentication.");
  }
});

// Set up a route to handle the Outlook authorization callback
app.get("/outlook/auth/callback", async (req, res) => {
  const code = req.query.code;

  try {
    // Exchange the Outlook authorization code for an access token
    const response = await msalClient.acquireTokenByCode({
      code,
      redirectUri: outlookRedirectUris[0],
      scopes: outlookScopes,
      clientSecret: outlookClientSecret, // Add the clientSecret parameter
    });

    // Store the Outlook access token in the session
    let outlookAccessToken = response.accessToken;
    const outlookToken = new Token({
      provider: "outlook",
      accessToken: outlookAccessToken,
    });
    await outlookToken.save();

    res.json({ accessToken: outlookAccessToken });
  } catch (error) {
    console.error("Error retrieving Outlook access token:", error);
    res.status(500).send("An error occurred during Outlook authentication.");
  }
});

app.get("/outlook/emails", async (req, res) => {
  try {
    let outlookAccessToken = await Token.find({ provider: "outlook" });
    console.log(outlookAccessToken[0], "outlookAccessToken");
    // Check if the user is authenticated with Outlook
    if (!outlookAccessToken[0]) {
      return res.status(401).send("User not authenticated with Outlook.");
    }

    // Create the Microsoft Graph client with the access token
    const client = Client.init({
      authProvider: (done) => {
        done(null, outlookAccessToken[0].accessToken);
      },
    });
    // Get the list of Outlook emails
    const response = await client.api("/me/mailfolders/inbox/messages").get();

    const emails = response.value;

    res.json(emails);
  } catch (error) {
    console.error("Error retrieving Outlook emails:", error);
    res.status(500).send("An error occurred while retrieving Outlook emails.");
  }
});

// Set up a route to get all sent Outlook emails
app.get("/outlook/sent/emails", async (req, res) => {
  try {
    // Check if the user is authenticated with Outlook
    let outlookAccessToken = await Token.find({ provider: "outlook" });
    console.log(outlookAccessToken[0], "outlookAccessToken");
    // Check if the user is authenticated with Outlook
    if (!outlookAccessToken[0]) {
      return res.status(401).send("User not authenticated with Outlook.");
    }

    // Create the Microsoft Graph client with the access token
    const client = Client.init({
      authProvider: (done) => {
        done(null, outlookAccessToken[0].accessToken);
      },
    });

    // Get the list of sent Outlook emails
    const response = await client
      .api("/me/mailfolders/sentitems/messages")
      .select("subject,from,toRecipients,body,attachments,receivedDateTime")
      .get();

    console.log(response); // Debugging: Log the response to check its structure

    const emails = response.value.map((email) => {
      const sender = email.from?.emailAddress?.address || "";
      const receivers = email.toRecipients.map((recipient) => recipient.emailAddress?.address);
      const subject = email.subject || "";
      const body = email.body?.content || "";
      const id = email?.id || "";

      const attachments =
        email.attachments?.map((attachment) => ({
          name: attachment.name,
          contentType: attachment.contentType,
        })) || [];
      const receivedDateTime = email.receivedDateTime || "";

      return {
        sender,
        receivers,
        subject,
        body,
        attachments,
        receivedDateTime,
        id,
      };
    });

    res.json(emails);
  } catch (error) {
    console.error("Error retrieving sent Outlook emails:", error);
    res.status(500).send("An error occurred while retrieving sent Outlook emails.");
  }
});

// Set up a route to get details for a specific Outlook email
app.get("/outlook/emails/:id", async (req, res) => {
  try {
    // Check if the user is authenticated with Outlook
    let outlookAccessToken = await Token.find({ provider: "outlook" });
    console.log(outlookAccessToken[0], "outlookAccessToken");
    // Check if the user is authenticated with Outlook
    if (!outlookAccessToken[0]) {
      return res.status(401).send("User not authenticated with Outlook.");
    }

    // Create the Microsoft Graph client with the access token
    const client = Client.init({
      authProvider: (done) => {
        done(null, outlookAccessToken[0].accessToken);
      },
    });
    let emailId = req.params.id;

    // Get the details for the specific Outlook email
    const response = await client
      .api(`/me/mailfolders/sentitems/messages/${emailId}`)
      .select("subject,from,toRecipients,body,attachments,receivedDateTime")
      .get();

    const email = response;

    const sender = email.from?.emailAddress?.address || "";
    const receivers = email.toRecipients.map((recipient) => recipient.emailAddress?.address);
    const subject = email.subject || "";
    const body = email.body?.content || "";
    const attachments =
      email.attachments?.map((attachment) => ({
        name: attachment.name,
        contentType: attachment.contentType,
      })) || [];
    const receivedDateTime = email.receivedDateTime || "";

    const emailDetails = {
      sender,
      receivers,
      subject,
      body,
      attachments,
      receivedDateTime,
    };

    res.json(emailDetails);
  } catch (error) {
    console.error("Error retrieving Outlook email details:", error);
    res.status(500).send("An error occurred while retrieving Outlook email details.");
  }
});

// Set up a route to get the list of Google Calendar events
app.get("/google/calendar/events", async (req, res) => {
  try {
    let gmailAccessToken = await Token.find({ provider: "gmail" });
    calendarOAuth2Client.setCredentials({ access_token: gmailAccessToken[0].accessToken });

    // Check if the user is authenticated with Google Calendar
    if (!gmailAccessToken[0].accessToken) {
      return res.status(401).send("User not authenticated with Google Calendar.");
    }

    // Create the Google Calendar API client with the access token
    const calendar = google.calendar({ version: "v3", auth: calendarOAuth2Client });

    // Get the list of Google Calendar events
    const response = await calendar.events.list({
      calendarId: "primary",
      timeMin: new Date().toISOString(), // Retrieve events from the current time onwards
      maxResults: 10, // Adjust the number of events to retrieve as per your requirement
      singleEvents: true,
      orderBy: "startTime",
    });

    const events = response.data.items;
    console.log(events, "this are the events of calender google");

    // Iterate over each event and extract the required details
    const eventDetails = events.map((event) => {
      const title = event.summary;
      const startTime = event.start.dateTime || event.start.date; // Use dateTime for events with specific start time, use date for all-day events
      const endTime = event.end.dateTime || event.end.date; // Use dateTime for events with specific end time, use date for all-day events
      const date = new Date(startTime).toLocaleDateString();

      return {
        title,
        startTime,
        endTime,
        date,
      };
    });

    res.json(eventDetails);
  } catch (error) {
    console.error("Error retrieving Google Calendar events:", error);
    res.status(500).send("An error occurred while retrieving Google Calendar events.");
  }
});

// Set up a route to get the list of Outlook calendar events
app.get("/outlook/calendar/events", async (req, res) => {
  try {
    // Check if the user is authenticated with Outlook
    let outlookAccessToken = await Token.find({ provider: "outlook" });
    console.log(outlookAccessToken[0], "outlookAccessToken");

    // Check if the user is authenticated with Outlook
    if (!outlookAccessToken[0]) {
      return res.status(401).send("User not authenticated with Outlook.");
    }

    // Create the Microsoft Graph client with the access token
    const client = Client.init({
      authProvider: (done) => {
        done(null, outlookAccessToken[0].accessToken);
      },
    });

    // Get the list of Outlook calendar events
    const response = await client.api("/me/events").get();

    const events = response.value;
    console.log(events, "this are the events of calender outlook");
    // Iterate over each event and retrieve the required details
    const eventDetails = events.map((event) => {
      const title = event.subject;
      const description = event.bodyPreview;
      const location = event.location.displayName;
      const startTime = event.start.dateTime;
      const endTime = event.end.dateTime;

      return {
        title,
        description,
        location,
        startTime,
        endTime,
      };
    });

    res.json(eventDetails);
  } catch (error) {
    console.error("Error retrieving Outlook calendar events:", error);
    res.status(500).send("An error occurred while retrieving Outlook calendar events.");
  }
});

// Set up a route to logout from Outlook
app.get("/outlook/logout", async (req, res) => {
  try {
    // Clear the Outlook access token from the database
    await Token.deleteOne({ provider: "outlook" });
    res.send("Logged out from Outlook.");
  } catch (error) {
    console.error("Error revoking Outlook access token:", error);
    res.status(500).send("An error occurred during Outlook logout.");
  }
});

// Set up a route to get all Google Calendar calendars
app.get("/google/calendar/calendars", async (req, res) => {
  try {
    let gmailAccessToken = await Token.findOne({ provider: "gmail" });
    calendarOAuth2Client.setCredentials({ access_token: gmailAccessToken.accessToken });

    // Check if the user is authenticated with Google Calendar
    if (!gmailAccessToken.accessToken) {
      return res.status(401).send("User not authenticated with Google Calendar.");
    }

    // Create the Google Calendar API client with the access token
    const calendar = google.calendar({ version: "v3", auth: calendarOAuth2Client });

    // Get the list of Google Calendar calendars
    const response = await calendar.calendarList.list();

    const calendars = response.data.items;

    res.json(calendars);
  } catch (error) {
    console.error("Error retrieving Google Calendar calendars:", error);
    res.status(500).send("An error occurred while retrieving Google Calendar calendars.");
  }
});

// Set up a route to get the list of Google Calendar events for a specific calendar
app.get("/google/calendar/:calendarId/events", async (req, res) => {
  try {
    let gmailAccessToken = await Token.findOne({ provider: "gmail" });
    calendarOAuth2Client.setCredentials({ access_token: gmailAccessToken.accessToken });

    // Check if the user is authenticated with Google Calendar
    if (!gmailAccessToken.accessToken) {
      return res.status(401).send("User not authenticated with Google Calendar.");
    }

    const calendarId = req.params.calendarId; // Get the calendar ID from the URL

    // Create the Google Calendar API client with the access token
    const calendar = google.calendar({ version: "v3", auth: calendarOAuth2Client });

    // Get the list of Google Calendar events for the specified calendar
    const response = await calendar.events.list({
      calendarId: calendarId,
      timeMin: new Date().toISOString(), // Retrieve events from the current time onwards
      maxResults: 1000, // Adjust get  all result
      singleEvents: true,
      orderBy: "startTime",
    });

    const events = response.data.items;

    // Iterate over each event and extract the required details
    const eventDetails = events.map((event) => {
      const title = event.summary;
      const startTime = event.start.dateTime || event.start.date; // Use dateTime for events with specific start time, use date for all-day events
      const endTime = event.end.dateTime || event.end.date; // Use dateTime for events with specific end time, use date for all-day events
      const date = new Date(startTime).toLocaleDateString();

      return {
        title,
        startTime,
        endTime,
        date,
      };
    });

    res.json(eventDetails);
  } catch (error) {
    console.error("Error retrieving Google Calendar events:", error);
    res.status(500).send("An error occurred while retrieving Google Calendar events.");
  }
});


// Start the server
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
