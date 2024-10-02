import asyncio
from io import StringIO
from azure.identity import UsernamePasswordCredential
from azure.identity import ClientSecretCredential
from kiota_abstractions.api_error import APIError
from kiota_authentication_azure.azure_identity_authentication_provider import AzureIdentityAuthenticationProvider
from msgraph import GraphRequestAdapter
from msgraph import GraphServiceClient
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.chat import Chat
from msgraph.generated.models.chat_type import ChatType
from msgraph.generated.models.aad_user_conversation_member import AadUserConversationMember
from msgraph.generated.models.user import User
import httpx
import giphy_client as gc
from giphy_client.rest import ApiException
import random as rd
from random import randint
from datetime import datetime
import numpy as np
import pandas as pd
import time
import configparser


#--------------------Import Config --------------------

#import logging #Enable for logging
#import http #Enable for HTTP requests review

#Enable HTTP logging
#----------------
#logging.basicConfig(
#    format="%(levelname)s [%(asctime)s] %(name)s - %(message)s",
#    datefmt="%Y-%m-%d %H:%M:%S",
#    level=logging.DEBUG
#)
#------------


#Enable HTTP logging
#----------------
#from http.client import HTTPConnection
#HTTPConnection.debuglevel = 1
#log = logging.getLogger('httpx')
#log.setLevel(logging.DEBUG)
#------------

# Review HTTP requests
#----------------
#def check_request():
#    old_send = http.client.HTTPConnection.send
#    def new_send(self, data):
#        print(f'{"-"*9} BEGIN REQUEST {"-"*9}')
#        print(data.decode('utf-8').strip())
#        print(f'{"-"*10} END REQUEST {"-"*10}')
#        return old_send(self, data)
#    http.client.HTTPConnection.send = new_send
#
#checkrequest()
#----------------


#--------------------Import Config --------------------

config = configparser.ConfigParser()
config.read('C:\Script\config.ini', encoding='utf-8')

ServiceAccountEmail = config.get('credentials', 'username')
#print(ServiceAccountEmail)


#--------------------------- Auth -------------------------------------------------
# Create authentication provider object.

#------Using Client Secret.  Required the following API permissions for App: Chat.Create, Teamwork.Migrate.All
#credential = ClientSecretCredential(
#    tenant_id=config.get('credentials', 'tenant_id'),
#    client_id=config.get('credentials', 'client_id'),
#    client_secret=config.get('credentials', 'client_secret'),
#)


#------------Using service account-------------------
credential = UsernamePasswordCredential(
    tenant_id=config.get('credentials', 'tenant_id'),
    client_id=config.get('credentials', 'client_id'),
    username=config.get('credentials', 'username'),
    password=config.get('credentials', 'password')
)

scopes = ['https://graph.microsoft.com/.default']
auth_provider = AzureIdentityAuthenticationProvider(credential, scopes=scopes)

# Initialize a request adapter with the auth provider.
request_adapter = GraphRequestAdapter(auth_provider)

# Create an API client with the request adapter.
msgraph = GraphServiceClient(request_adapter)

# Create an event loop
loop = asyncio.get_event_loop()

#--------------------------- /Auth ------------------------------------------------


#--------------------------- Functions  -------------------------------------------
#------------------------------
async def get_me(): # Test Get Service Account user name
    try:
        user = await msgraph.me.get()
        if user:
            print(user.user_principal_name, user.display_name, user.id)
            print("MSGraph Auth is working !!!")
    except APIError as e:
        print(f'Error: {e.error.message}')


#------------------------------
async def getUserIdByEmail(UserEmail): # Get Azure UserID by User Email 
    try:
        #print("######## Run getUserIdByEmail", UserEmail)
        user = User()
        user = await msgraph.users.by_user_id(UserEmail).get()
        #print ("##################Finish",UserEmail)
        if user:
            print(user.user_principal_name, user.display_name, user.id)
            UserID = user.id
            return(UserID)
    except APIError as e:
        print(f'Error: {e.error.message}')


#------------------------------
async def getGroupMembers(): # Get a list of users from AAD group by Group ID 
    try:
        members = await msgraph.groups.by_group_id(GroupID).members.get()
        group = await msgraph.groups.by_group_id(GroupID).get()
        print(group.display_name)
        if members and members.value:
            for member in members.value:
                user = await msgraph.users.by_user_id(member.id).get()
                if user:
                    print(user.display_name, user.mail, user.id)
    except APIError as e:
        print(e.error.message)


#------------------------------
async def getChats(UserEmail): # Recieve a list of existing chats for yourself and print ChatID and members
    try:
        chats = await msgraph.users.by_user_id(UserEmail).chats.get()
        if chats and chats.value:
            for chat in chats.value:
                chatID = await msgraph.users.by_user_id(UserEmail).chats.by_chat_id(chat.id).get()
                if chatID:
                    print("----------------------------")
                    print("ChatID = ", chatID.id)
                    ChatMembers = await msgraph.users.by_user_id(UserEmail).chats.by_chat_id(chat.id).members.get()
                    for ChatMember in ChatMembers.value:
                        print(ChatMember.display_name, ChatMember.roles)
                                  
    except APIError as e:
        print(e.error.message)


#------------------------------
async def sendChatMessage(CreatedChatID, ChatMessageContent):  # Send a message (ChatMessageContecnt) to User2 from CreatedChatID)
    try:
        #print("######## Run sendChatMessage with CreatedChatID" ,CreatedChatID, "and ChatMessageContent", ChatMessageContent)
        Message = ChatMessage()
        MessageBody = ItemBody()
        MessageBody.content = ChatMessageContent # Message content
        MessageBody.content_type = BodyType.Html
        Message.body = MessageBody
        await msgraph.chats.by_chat_id(CreatedChatID).messages.post(Message)
        #print("###### Message was setn with Body:", Message.body.content)
    except APIError as e:
        print(e.error.message)


#------------------------------
async def createGifChatMessage(CreatedChatID, GifContent):  # Create a message with GIF to User2 from CreatedChatID)
    try:
        Message = ChatMessage()
        MessageBody = ItemBody()
        MessageBody.content = '<html><body><img height="297" src="{}" width="297" style="vertical-align:bottom; width:297px; height:297px">'.format(GifContent)
        MessageBody.content_type = BodyType.Html
        Message.body = MessageBody
        #print("#############  RUN Send Message with ChatID", CreatedChatID)
        await msgraph.chats.by_chat_id(CreatedChatID).messages.post(Message)
        #print(Message.body.content)
    except APIError as e:
        print(e.error.message)

#------------------------------
async def createNewChat(UserID1, UserID2): # Create a new chat between Service Account/ChatBot (User1) and EndUser (User2)
    try:
        #print("######## RUN createNewChat with UserIDs", UserID1, UserID2)
        #Create Member 1 (Service Account) for 1-1 chat.
        User1 = AadUserConversationMember()
        User1.user_id = UserID1 #UserID of ServiceAccount
        User1.roles = ["owner"]
        User1.odata_type = "#microsoft.graph.aadUserConversationMember"
        User1.additional_data = {"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + UserID1 +"')"}
        
        #Create Member 2 (EndUser) for 1-1 chat.
        User2 = AadUserConversationMember()
        User2.user_id = UserID2 #UserID - recipient
        User2.roles = ["owner"]
        User2.odata_type = "#microsoft.graph.aadUserConversationMember"
        User2.additional_data = {"user@odata.bind": "https://graph.microsoft.com/v1.0/users('" + UserID2 +"')"}
        #print (User1.additional_data)
        #print (User2.additional_data)
        NewChatBody = Chat()
        NewChatBody.members = (User1,User2)
        NewChatBody.chat_type = ChatType.OneOnOne

        #Create a new Chat and get Chat ID
        CreatedChat = await msgraph.chats.post(body=NewChatBody)
        CreatedChatID = CreatedChat.id
        #print("------------------------")
        #print("Chat Created", CreatedChatID)
        #print("------------------------")
        return(CreatedChatID)
    except APIError as e:
        print(e.error.message)


#------------------------------
async def getChatMembers(ChatID): #Get a list of members from Chat by ChatID. Not in use by default.
    try:
        members = await msgraph.chats.by_chat_id(ChatID).members.get()
        if members and members.value:
            for member in members.value:
                    print(member.additional_data)
    except APIError as e:
        print(e.error.message)


#------------------------------
async def createMSTeamsChatAndSendMessage(UserEmail, ChatMessageContent): # Create a new chat by UserEmail and send a message with ChatMessageContent
    #print("######  RUN createMSTeamsChatAndSendMessage with UserEmail, ChatMessageContent", UserEmail, ChatMessageContent)
    UserID2 = await getUserIdByEmail(UserEmail)
    ChatID = await createNewChat(UserID1, UserID2)
    #print("#######  Back to CreateMSTeamsChatAndSendMessage with ChatID",ChatID)
    await sendChatMessage(ChatID, ChatMessageContent)


#------------------------------
async def createMSTeamsChatAndMessageWGif(UserEmail, search_term): #  Create a new chat by UserEmail and send a GIF based on user properties
    UserID2 = await getUserIdByEmail(UserEmail)
    ChatID = await createNewChat(UserID1, UserID2)
    #print("#######  Back to CreateMSTeamsChatAndMessageWGif with ChatID",ChatID)
    GifContent = await createGifUrl(search_term)
    await createGifChatMessage(ChatID, GifContent)


#------------------------------
async def createGifUrl(search_term): #Create URL to Gif based on User search_term
    api_instance = gc.DefaultApi()

    gif_api_key = config.get('credentials', 'gif_api_key')
    query = search_term
    fmt = 'gif'

    try:
        response = api_instance.gifs_search_get(gif_api_key,query,limit=25,offset=randint(1,25),fmt=fmt)
        gif_id = response.data[0]
        url_gif = gif_id.images.downsized.url
        #print("####### URL to GIF=", url_gif)
        return (url_gif)
    except ApiException as e:
        print("Exception when calling DefaultApi->gifs_search_get: %s\n" % e)


#------------------------------
def getSnowReport(url): # Get SNOW Report as CSV
    headers = {"Accept":"application/json"}
    SnowUser = config.get('credentials', 'SnowUser')
    SnowPWD = config.get('credentials', 'SnowPWD')
    response = httpx.get(url=url, auth=(SnowUser, SnowPWD), headers=headers, timeout=30)
    response.encoding = 'utf-8'
    print(url)

    if response.status_code == 200:
        df = pd.read_csv(StringIO(response.text))
        return(df)
    else:
        print(f"Request failed with code {response.status_code}")


#------------------------------
def BuildSNOWReport(): # Build Snow report for unassigned tickets and and out of SLA (use SLADays as config for SLA)
    #---------------------Servicenow Tasks (from RITM) ---------------------------------
    
    SnowReportTask = config.get('SNOW', 'SnowReportTask')
    tickets = getSnowReport(url=f"https://aligntech.service-now.com/sys_report_template.do?CSV&jvar_report_id={SnowReportTask}")
    tickets["assigned_to"].fillna("", inplace=True)
    tickets["aging_days"] = tickets["opened_at"].apply(lambda x: np.busday_count(np.datetime64(datetime.strptime(x, "%m/%d/%y %H:%M:%S")).astype("M8[D]"), np.datetime64("today")))
    tickets["number"] = tickets["number"].apply(lambda x: makeSnowTaskUrl(x))
    
    ticketsShort = tickets[["number", "state", "assigned_to", "assigned_to.email", "request_item.request.requested_for", "request_item.request.requested_for.u_office_string", "opened_at", "aging_days", "short_description"]].copy()
    ticketsShort.columns = ["Ticket", "State", "Agent", "Agent Email", "Requestor", "Office", "Opened", "Aging", "Description"]
  
    #--------------Incidents--------------------------------------------------------------
    SnowReportInc = config.get('SNOW', 'SnowReportInc')
    incidents = getSnowReport(url=f"https://aligntech.service-now.com/sys_report_template.do?CSV&jvar_report_id={SnowReportInc}")
    incidents["assigned_to"].fillna("", inplace=True)
    incidents.rename(columns={"sys_created_on": "opened_at"}, inplace=True)     

    incidents["aging_days"] = incidents["opened_at"].apply(lambda x: np.busday_count(np.datetime64(datetime.strptime(x, "%m/%d/%y %H:%M:%S")).astype("M8[D]"), np.datetime64("today")))
    incidents["number"] = incidents["number"].apply(lambda x: makeSnowTaskUrl(x))

    incidentsShort = incidents[["number", "state", "assigned_to", "assigned_to.email", "caller_id", "caller_id.u_office_string", "opened_at", "aging_days", "short_description"]].copy()
    incidentsShort.columns = ["Ticket", "State", "Agent", "Agent Email", "Requestor", "Office", "Opened", "Aging", "Description"]

    combined = pd.concat([incidentsShort, ticketsShort])

    #--------------- Preparation of summary for tickets
    unassignedTickets = combined.loc[(combined["State"].isin (["New", "In Progress", "On Hold", "Open", "Work in Progress", "Pending"])) & (combined["Agent"]=="")][["Ticket", "Requestor", "State", "Opened", "Aging", "Description"]].copy()
    unassignedTickets.columns = ["Ticket", "Requestor", "State", "Opened", "Aging", "Description"]
    unassignedTickets.reset_index(drop=True, inplace=True)

    outOfSLATickets = combined.loc[(combined["Aging"] > SLADays) & (combined["State"].isin(["New", "In Progress", "On Hold", "Open", "Work in Progress", "Pending"]))][["Ticket", "Agent", "Requestor", "State", "Opened", "Aging", "Description"]].copy()
    outOfSLATickets.columns = ["Ticket", "Assigned to", "Requestor", "State", "Opened", "Aging", "Description"]
    outOfSLATickets.reset_index(drop=True, inplace=True)

    return(combined, unassignedTickets, outOfSLATickets)


#------------------------------
def makeSnowTaskUrl(ticket): #Create a hyperlink to ticket in SNOW
    return(f'<a href="https://aligntech.service-now.com/nav_to.do?uri=task.do?sysparm_query=number={ticket}">{ticket}</a>')


#------------------------------
def getGifSearchTerm(agentEmail): #Get Gif search term based on predefined Dict----
    if agentEmail in gifsDict:
        gifs = gifsDict[agentEmail].split(',')
        gif = gifs[rd.randint(0, len(gifs) - 1)]
        #print("Gif=",gif)
    else:
        gif = "happy"
        #print("GifNotFound and = Happy")

    return(gif)


#------------------------------ 
def getGreeting(agentEmail, agent): #Get Greeting based on spoken language----
    if agentEmail in languagesSpokenDict:
        languages = languagesSpokenDict[agentEmail].split(',')
        lang = languages[rd.randint(0, len(languages) - 1)]
    else:
        lang = "en"

    greetings = greatingsDict[lang].split('(*)')
    greet = greetings[rd.randint(0, len(greetings) - 1)]

    return(greet.format(agent))


#---------------- 
async def SendWeeklyReminder(GroupID,ReminderMessage): #Send reminder to all member of GroupID every Friday (5)
   
    if datetime.today().isoweekday() == 5:

        try:
            members = await msgraph.groups.by_group_id(GroupID).members.get()
            #group = await msgraph.groups.by_group_id(GroupID).get()
            #print(group.display_name)
            
            if members and members.value:
                
                for member in members.value:
                    user = await msgraph.users.by_user_id(member.id).get()
                    
                    if user:
                        #print(user.display_name, user.mail, user.id)
                        #user.mail = "TestEmail@aligntech.com"
                        await createMSTeamsChatAndSendMessage(user.mail, ReminderMessage)
                        await createMSTeamsChatAndMessageWGif(user.mail, 'reminder')

        except APIError as e:
            print(e.error.message)
       
    else:
        print('Today is not Friday')
        

#----------------
async def SendSummaryToSupervisors(unassignedTickets,outOfSLATickets): #Send summary of Unassigned and Oot of SLA tickets to Supervisor 
    
    for sup in supervisors:
        supervisorName = sup.capitalize()
        supervisorEmail = supervisors[sup]

        greeting = getGreeting(supervisorEmail, supervisorName)

        await createMSTeamsChatAndSendMessage(supervisorEmail, greeting)

        if len(unassignedTickets) != 0:
            await createMSTeamsChatAndSendMessage(supervisorEmail, "Please review the following UNASSIGNED tickets:")
            await createMSTeamsChatAndSendMessage(supervisorEmail, unassignedTickets.to_html(escape=False))

        if len(outOfSLATickets) != 0:
            await createMSTeamsChatAndSendMessage(supervisorEmail, "The tickets that are falling outside of SLA:")
            await createMSTeamsChatAndSendMessage(supervisorEmail, outOfSLATickets.to_html(escape=False))

        await createMSTeamsChatAndMessageWGif(supervisorEmail, getGifSearchTerm(supervisorEmail))


#---------------- 
async def SendSummaryToAgent(combined): #Send to Agent a list of Out of SLA tickets assigned to him
    agentsEmails = combined["Agent Email"].unique()
    agentsGrouped = pd.DataFrame(combined.groupby(by = ["Agent Email", "Agent"], as_index = False).count())

    for agentEmail in agentsEmails:
        
        if agentEmail in supervisorEmails:
            continue
        
        agentOutOfSLATickets = combined.loc[(combined["Agent Email"] == agentEmail) & (combined["Aging"] > SLADays) & (combined["State"].isin(["New", "In Progress", "Open", "Pending", "Work in Progress", "On Hold"]))][["Ticket", "Agent", "Requestor", "State", "Opened", "Aging", "Description"]].copy()

        if len(agentOutOfSLATickets) != 0:
            agentOutOfSLATicketsShort = agentOutOfSLATickets[["Ticket", "Requestor", "State", "Opened", "Aging", "Description"]].copy()
            agentOutOfSLATicketsShort.reset_index(drop=True, inplace=True)
            
            agentName = agentsGrouped.loc[(agentsGrouped["Agent Email"] == agentEmail)].Agent.values[0]
            greeting = getGreeting(agentEmail, agentName)
            #print(greeting)        

            #display(HTML(agentOutOfSLATicketsShort.to_html(escape=False)))
            #agentEmail = "TestEmail@aligntech.com" #Send Agent results to yourself

            await createMSTeamsChatAndSendMessage(agentEmail, greeting)
            await createMSTeamsChatAndSendMessage(agentEmail, agentOutOfSLATicketsShort.to_html(escape=False))

            await createMSTeamsChatAndMessageWGif(agentEmail, getGifSearchTerm(agentEmail))
            
            time.sleep(rd.randint(5, 10))


#---------------- 
async def MainRun(): #Run Main Function from Mon to Fri (days 1-5)
    if datetime.today().isoweekday() in [1,2,3,4,5]:
       combined, unassignedTickets, outOfSLATickets = BuildSNOWReport()
       await SendSummaryToSupervisors(unassignedTickets,outOfSLATickets)
       await SendSummaryToAgent(combined)
       await SendWeeklyReminder(GroupID, ReminderMessage) # Send weekly reminder
    else:
        print('Weekend')

# ---------------------/Functions-------------------------------------------------


# ---------------------Properties --------------------------------------------


#------------------------------languagesSpokenDict----------
#--------- You can use build-in languagesSpokenDict instead of config.ini
#languagesSpokenDict = {
#    "user1@aligntech.com": ["ru"],
#    "user2@aligntech.com": ["pl", "en"],
#}
languagesSpokenDict = dict(config.items('languagesSpokenDict'))


#------------------------------gifsDict-------------- 
#--------- You can use build-in gifsDict instead of config.ini
#gifsDict = {
#    "user1@aligntech.com": ["please", "quickly", "fail"],
#    "user2@aligntech.com": ["please", "quickly", "bless"]
#}
gifsDict = dict(config.items('gifsDict'))

#------------------------------greatingsDict-------------
#-----You can use build-in greatingsDict instead of config.ini
#greatingsDict = {
#    "ru": ["Доброе утро, комрад {}! Посмотри, пожалуйста, тикеты.", "Привет, братюнь, {}! Что-то тикеты копятся, посмотри по-братски."],
#    "en": ["Good morning, my friend, {}! Please find a moment to review your tickets.", "Howdy {}? Tickets are piling up :( take a look, please."],
#    "de": ["Guten morgen, Parteigenosse {}! Bitte überprüfen Sie Ihre Tickets, unsere Benutzer leiden :("],
#    "pt": ["Bom dia {}! Por favor, revise seus tickets, nossos usuários estão sofrendo :("],
#    "es": ["Buenos días mi amigo {}! Encuentre un momento para revisar sus boletos."],
#    "pl": ["Dzień dobry towarzyszu {}! Proszę spojrzeć na kolejkę żądań.", "Cześć towarzyszu {}! Niektóre zgłoszenia piętrzą się , rzuć okiem na nie!"]
#}
greatingsDict = dict(config.items('greatingsDict'))


#------The list of supervisors for SendSummaryToSupervisors()-------------
#----- You can use build-in supervisors instead of config.ini
#supervisors = {
#    "User1": "user1@aligntech.com",
#    "User2": "user2@aligntech.com"
#    }
supervisors = dict(config.items('Supervisors'))


#---------- The list of supervisors for SendSummaryToAgent() to exclude from agent list-------
#---------You can use build-in greatingsDict instead of config.ini
#supervisorEmails = ["User1@aligntech.com", "User2@aligntech.com"]supervisorEmails = list(supervisors.values())
supervisorEmails = list(supervisors.values())


#---------------- Others config

#SLA for BuildSNOWReport()
#SLADays = 3 
SLADays = config.getint('Others', 'SLADays')

#Reminder message for Weekly Reminders
#ReminderMessage = "Dont forget about weekly report"
ReminderMessage = config.get('Others', 'ReminderMessage')

#Group ID for weekly reminder
#GroupID = "aasdasdasdsadasdasdasd"
GroupID = config.get('Others', 'GroupID')


# Get UserID for service account
UserID1 = loop.run_until_complete(getUserIdByEmail(ServiceAccountEmail)) # Get UserID for service account.

#----------------------/Properties --------------------------------------------



#------------Test section------------------
# Run your test async functions using the same event loop
# Uncomment lines with "TestEmail@aligntech.com" and replace to your email to get all messages


#loop.run_until_complete(get_me())
#loop.run_until_complete(createMSTeamsChatAndSendMessage("TestEmail@aligntech.com", "Test Message"))
#loop.run_until_complete(createMSTeamsChatAndMessageWGif("TestEmail@aligntech.com", getGifSearchTerm("test")))
#loop.run_until_complete(get_me())

#print("UserID =",loop.run_until_complete(getUserIdByEmail("TestUser@aligntech.com")))

#------------/Test section------------------

#------------Run functions section------------------

loop.run_until_complete(MainRun()) # Uncomment to RUN

# -------------------- Close the event loop---------------------------
loop.close()
