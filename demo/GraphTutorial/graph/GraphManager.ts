//  Copyright (c) Microsoft. All rights reserved.
//  Licensed under the MIT license.

import { Client } from '@microsoft/microsoft-graph-client';

import { GraphAuthProvider } from './GraphAuthProvider';
import { Alert } from 'react-native';
// Set the authProvider to an instance
// of GraphAuthProvider
const clientOptions = {
  authProvider: new GraphAuthProvider()
};

// Initialize the client
const graphClient = Client.initWithMiddleware(clientOptions);
//const graphClient = Client.init(clientOptions);

export class GraphManager {
  static getUserAsync = async () => {
    // GET /me
    return graphClient.api('/me').get();
  }

  // <GetEventsSnippet>
  static getEvents = async () => {
    // GET /me/events
    return graphClient.api('/me/events')
      // $select='subject,organizer,start,end'
      // Only return these fields in results
      .select('subject,organizer,start,end')
      // $orderby=createdDateTime DESC
      // Sort results by when they were created, newest first
      .orderby('createdDateTime DESC')
      .get();
    /* let res = await graphClient.api(' /users/Satheeshkumar.S@coats.com/events')
      .header('Prefer', 'outlook.timezone="Pacific Standard Time"')
      .select('subject,body,bodyPreview,organizer,attendees,start,end,location')
      .get(); 
    return res*/
  }
  // </GetEventsSnippet>

  static getOtherEventsByme = async () => {
    // GET /me/events
    return graphClient.api('/me/findMeetingTimes')
      // $select='subject,organizer,start,end'
      // Only return these fields in results
      .select('subject,organizer,start,end')
      // $orderby=createdDateTime DESC
      // Sort results by when they were created, newest first
      .orderby('createdDateTime DESC')
      .get();
  }
  static getMeetingTimeSuggestions = async () => {
    // GET /users/events/findMeetingTimes
    const meetingTimeSuggestionsResult = {
      attendees: [
        /*   {
            type: "required",
            emailAddress: {
              name: "Satheeshkumar S",
              address: "Satheeshkumar.S@coats.com"
            }
          }, */
        {
          type: "required",
          emailAddress: {
            name: "Linjith NP",
            address: "Linjith.NP@coats.com"
          }
        }
      ],
      locationConstraint: {
        isRequired: "false",
        suggestLocation: "false",
        locations: [
          {
            resolveAvailability: "false",
            displayName: "Conf room Hood"
          }
        ]
      },
      timeConstraint: {
        activityDomain: "work",
        timeSlots: [
          {
            start: {
              dateTime: "2020-10-15T09:00:00",
              timeZone: "UTC"
            },
            end: {
              dateTime: "2020-10-16T17:00:00",
              timeZone: "UTC"
            }
          }
        ]
      },
      isOrganizerOptional: "false",
      meetingDuration: "PT1H",
      returnSuggestionReasons: "true",
      minimumAttendeePercentage: "100"
    };
    // let res = await graphClient.api('/users/Linjith.NP@coats.com/findMeetingTimes').post(meetingTimeSuggestionsResult);
    let res = await graphClient.api('/me/findMeetingTimes').post(meetingTimeSuggestionsResult);
    console.log(`meetings--->${JSON.stringify(res)}`)
    Alert.alert('MeetingsRes: ' + JSON.stringify(res));
    return res
  }
  static getSharedEvents = async () => {
    // GET /users/events
    return graphClient.api('/users/Linjith.NP@coats.com/calendar/events')
      // $select='subject,organizer,start,end'
      // Only return these fields in results
      .select('subject,organizer,start,end')
      // $orderby=createdDateTime DESC
      // Sort results by when they were created, newest first
      .orderby('createdDateTime DESC')
      .get();
  }

  static getSchedule = async () => {
    // GET /users/events
    // graphClient.api('/users/Linjith.NP@coats.com/calendar/events')
    const scheduleInformation = {
      schedules: ["Satheeshkumar.S@coats.com", "Linjith.NP@coats.com"],
      startTime: {
        dateTime: "2020-10-15T09:00:00",
        timeZone: "UTC"
      },
      endTime: {
        dateTime: "2020-10-16T18:00:00",
        timeZone: "UTC"
      },
      availabilityViewInterval: 60
    };

    let res = await graphClient.api('/me/calendar/getSchedule')
      .post(scheduleInformation);
    console.log("/me/calendar/getSchedule======>", JSON.stringify(res.value[0].scheduleItems))
    return res
  }

  static scheduleMeeting = async () => {
    // POST /users/events
    const event = {
      subject: "Let's go for lunch",
      body: {
        contentType: "HTML",
        content: "Does noon work for you?"
      },
      start: {
        dateTime: "2020-10-15T12:00:00",
        timeZone: "Pacific Standard Time"
      },
      end: {
        dateTime: "2020-10-15T14:00:00",
        timeZone: "Pacific Standard Time"
      },
      location: {
        displayName: "Harry's Bar"
      },
      attendees: [
        {
          emailAddress: {
            address: "Linjith.NP@coats.com",
            name: "Linjith NP"
          },
          type: "required"
        },
        {
          emailAddress: {
            address: "Satheeshkumar.S@coats.com",
            name: "Satheesh"
          },
          type: "required"
        },
        {
          emailAddress: {
            address: "Santhiya.Jawahar@coats.com",
            name: "Santhiya"
          },
          type: "optional"
        }

      ],
      allowNewTimeProposals: true,
      transactionId: "7E163156-7762-4BEB-A1C6-729EA81755A6"
    };
    let res = await graphClient.api('/me/events').post(event);
    return res
  }

}
