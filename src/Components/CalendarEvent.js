import React, { useState, useEffect } from 'react';
import MicrosoftGraphService from '../Services/MicrosoftGraphService';
import { PublicClientApplication } from '@azure/msal-browser';

const msalConfig = {
  auth: {
    clientId: 'YOUR_CLIENT_ID_HERE', // Replace with your Azure AD application's client ID
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID_HERE', // Replace with your Azure AD tenant ID
    redirectUri: 'http://localhost:3000/', // Ensure this matches the registered redirect URI
  },
};

const CalendarEvent = () => {
  const [events, setEvents] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState(null);
  const [msalInstance, setMsalInstance] = useState(null);

  const fetchData = async () => {
    if (!msalInstance) {
      return;
    }

    try {
      const accessToken = await msalInstance.acquireTokenSilent({
        scopes: ['openid', 'profile', 'User.Read', 'Calendars.Read'],
      });

      const userObjectId = await MicrosoftGraphService.getUserObjectId('user@example.com');
      const newEvents = await MicrosoftGraphService.getTeamsCalendarEventsForUser(userObjectId);

      setEvents(newEvents);
    } catch (err) {
      setError(err.message || 'An error occurred while fetching events');
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    async function initializeMsal() {
      try {
        const instance = new PublicClientApplication(msalConfig);
        await instance.handleRedirectPromise();
        setMsalInstance(instance);
      } catch (err) {
        setError(err);
      } finally {
        setIsLoading(false);
      }
    }

    initializeMsal();
  }, []);

  useEffect(() => {
    fetchData();

    const intervalId = setInterval(fetchData, 30000);

    return () => clearInterval(intervalId);
  }, [msalInstance]);

  if (isLoading) {
    return <div>Loading...</div>;
  }

  if (error) {
    return <div>Error: {error.message}</div>;
  }

  return (
    <div>
      <h2>Calendar Events</h2>
      <ul>
        {events.map((event, index) => (
          <li key={index}>
            <div>Title: {event.subject}</div>
            <div>Start Time: {new Date(event.start.dateTime).toLocaleString()}</div>
            <div>End Time: {new Date(event.end.dateTime).toLocaleString()}</div>
            <div>Meeting Link: <a href={event.onlineMeeting.joinUrl} target="_blank" rel="noopener noreferrer">Open meeting</a></div>
            <div>Organizer: {event.organizer.emailAddress.name}</div>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default CalendarEvent;
