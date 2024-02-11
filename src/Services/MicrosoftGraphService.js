import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken } from './getAccessToken';

const MicrosoftGraphService = {
  getUserObjectId: async (email) => {
    try {
      const accessToken = await getAccessToken();

      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });


      const user = await client.api(`/users/${email}`).get();
      return user.id;
    } catch (error) {
      console.error('Error fetching user object ID:', error);
      throw error;
    }
  },

  getTeamsCalendarEventsForUser: async (email) => {
    try {
      const userId = await MicrosoftGraphService.getUserObjectId(email);
      const accessToken = await getAccessToken();
  
      console.log('User ID:', userId);
      console.log('Access Token:', accessToken);
  
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });
  
 
      const events = await client.api(`/users/${userId}/calendar/events`).get();
  
      return events.value;
    } catch (error) {
      console.error('Error fetching calendar events for Microsoft Teams user:', error);
      throw error;
    }
  },
}
  

export default MicrosoftGraphService;
