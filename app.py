import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import List, Optional
import os
from datetime import datetime, date
from enum import Enum
import requests

# Page configuration
st.set_page_config(
    page_title="EventOps — Event Management Software",
    page_icon="🎉",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .hero-text {
        font-size: 1.5rem;
        color: #333;
        margin-bottom: 2rem;
    }
    .feature-card {
        padding: 1.5rem;
        border-radius: 10px;
        background: #f8f9fa;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #1f77b4;
        color: white;
    }
</style>
""", unsafe_allow_html=True)


# Hugging Face API Handler
class HuggingFaceHandler:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.api_url = "https://api-inference.huggingface.co/models"
    
    def generate_text(self, model: str, prompt: str, max_length: int = 500):
        """Generate text using Hugging Face API"""
        try:
            headers = {"Authorization": f"Bearer {self.api_key}"}
            payload = {
                "inputs": prompt,
                "parameters": {"max_length": max_length, "temperature": 0.7}
            }
            response = requests.post(
                f"{self.api_url}/{model}",
                headers=headers,
                json=payload
            )
            if response.status_code == 200:
                result = response.json()
                if isinstance(result, list) and len(result) > 0:
                    return result[0].get("generated_text", prompt)
                return result.get("generated_text", prompt)
            return f"Error: {response.status_code}"
        except Exception as e:
            return f"Error: {str(e)}"


# Excel Handler
class ExcelHandler:
    def __init__(self, file_path: str):
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        self.file_path = file_path
        self.ensure_file_exists()
    
    def ensure_file_exists(self):
        if not os.path.exists(self.file_path):
            from openpyxl import Workbook
            wb = Workbook()
            wb.remove(wb.active)
            sheets = [
                'clients', 'venues', 'vendors', 'events', 'sub_events',
                'proposals', 'budget_items', 'tasks', 'runsheet', 'team_members',
                'creative_directions'
            ]
            for sheet_name in sheets:
                wb.create_sheet(sheet_name)
            wb.save(self.file_path)
    
    def _get_sheet_data(self, sheet_name: str) -> pd.DataFrame:
        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            return df
        except:
            return pd.DataFrame()
    
    def _save_sheet_data(self, sheet_name: str, data: List[dict]):
        df = pd.DataFrame(data)
        wb = load_workbook(self.file_path)
        if sheet_name in wb.sheetnames:
            wb.remove(wb[sheet_name])
        ws = wb.create_sheet(sheet_name)
        if not df.empty:
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
        wb.save(self.file_path)
    
    def _get_next_id(self, sheet_name: str) -> int:
        df = self._get_sheet_data(sheet_name)
        if df.empty or 'id' not in df.columns:
            return 1
        return df['id'].max() + 1 if df['id'].max() else 1
    
    def get_clients(self) -> List[dict]:
        df = self._get_sheet_data('clients')
        return df.to_dict('records')
    
    def create_client(self, client_data: dict) -> dict:
        client_data['id'] = self._get_next_id('clients')
        client_data['created_at'] = datetime.now().isoformat()
        clients = self.get_clients()
        clients.append(client_data)
        self._save_sheet_data('clients', clients)
        return client_data
    
    def get_venues(self) -> List[dict]:
        df = self._get_sheet_data('venues')
        return df.to_dict('records')
    
    def create_venue(self, venue_data: dict) -> dict:
        venue_data['id'] = self._get_next_id('venues')
        venues = self.get_venues()
        venues.append(venue_data)
        self._save_sheet_data('venues', venues)
        return venue_data
    
    def get_vendors(self) -> List[dict]:
        df = self._get_sheet_data('vendors')
        return df.to_dict('records')
    
    def create_vendor(self, vendor_data: dict) -> dict:
        vendor_data['id'] = self._get_next_id('vendors')
        vendors = self.get_vendors()
        vendors.append(vendor_data)
        self._save_sheet_data('vendors', vendors)
        return vendor_data
    
    def get_events(self) -> List[dict]:
        df = self._get_sheet_data('events')
        return df.to_dict('records')
    
    def create_event(self, event_data: dict) -> dict:
        event_data['id'] = self._get_next_id('events')
        event_data['created_at'] = datetime.now().isoformat()
        event_data['updated_at'] = datetime.now().isoformat()
        events = self.get_events()
        events.append(event_data)
        self._save_sheet_data('events', events)
        return event_data
    
    def get_sub_events(self, event_id: int) -> List[dict]:
        df = self._get_sheet_data('sub_events')
        items = df.to_dict('records')
        return [item for item in items if item.get('event_id') == event_id]
    
    def create_sub_event(self, item_data: dict) -> dict:
        item_data['id'] = self._get_next_id('sub_events')
        items = self._get_sheet_data('sub_events').to_dict('records')
        items.append(item_data)
        self._save_sheet_data('sub_events', items)
        return item_data
    
    def get_budget_items(self, event_id: int) -> List[dict]:
        df = self._get_sheet_data('budget_items')
        items = df.to_dict('records')
        return [item for item in items if item.get('event_id') == event_id]
    
    def create_budget_item(self, item_data: dict) -> dict:
        item_data['id'] = self._get_next_id('budget_items')
        items = self._get_sheet_data('budget_items').to_dict('records')
        items.append(item_data)
        self._save_sheet_data('budget_items', items)
        return item_data
    
    def get_tasks(self, event_id: int) -> List[dict]:
        df = self._get_sheet_data('tasks')
        tasks = df.to_dict('records')
        return [task for task in tasks if task.get('event_id') == event_id]
    
    def create_task(self, task_data: dict) -> dict:
        task_data['id'] = self._get_next_id('tasks')
        task_data['created_at'] = datetime.now().isoformat()
        tasks = self._get_sheet_data('tasks').to_dict('records')
        tasks.append(task_data)
        self._save_sheet_data('tasks', tasks)
        return task_data
    
    def get_runsheet(self, event_id: int) -> List[dict]:
        df = self._get_sheet_data('runsheet')
        items = df.to_dict('records')
        return [item for item in items if item.get('event_id') == event_id]
    
    def create_runsheet_item(self, item_data: dict) -> dict:
        item_data['id'] = self._get_next_id('runsheet')
        items = self._get_sheet_data('runsheet').to_dict('records')
        items.append(item_data)
        self._save_sheet_data('runsheet', items)
        return item_data
    
    def get_team_members(self) -> List[dict]:
        df = self._get_sheet_data('team_members')
        return df.to_dict('records')
    
    def create_team_member(self, member_data: dict) -> dict:
        member_data['id'] = self._get_next_id('team_members')
        members = self.get_team_members()
        members.append(member_data)
        self._save_sheet_data('team_members', members)
        return member_data
    
    def get_creative_directions(self, event_id: int) -> List[dict]:
        df = self._get_sheet_data('creative_directions')
        items = df.to_dict('records')
        return [item for item in items if item.get('event_id') == event_id]
    
    def create_creative_direction(self, direction_data: dict) -> dict:
        direction_data['id'] = self._get_next_id('creative_directions')
        direction_data['created_at'] = datetime.now().isoformat()
        items = self._get_sheet_data('creative_directions').to_dict('records')
        items.append(direction_data)
        self._save_sheet_data('creative_directions', items)
        return direction_data


# Initialize Excel handler
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
EXCEL_FILE = os.path.join(DATA_DIR, "eventops_data.xlsx")
excel_handler = ExcelHandler(EXCEL_FILE)

# Initialize session state
if 'clients' not in st.session_state:
    st.session_state.clients = excel_handler.get_clients()
if 'venues' not in st.session_state:
    st.session_state.venues = excel_handler.get_venues()
if 'vendors' not in st.session_state:
    st.session_state.vendors = excel_handler.get_vendors()
if 'events' not in st.session_state:
    st.session_state.events = excel_handler.get_events()
if 'team_members' not in st.session_state:
    st.session_state.team_members = excel_handler.get_team_members()


# Sidebar
st.sidebar.title("EventOps 🎉")
st.sidebar.markdown("---")

# Hugging Face API Key
hf_api_key = st.sidebar.text_input("Hugging Face API Key", type="password", help="Get your free API key from huggingface.co/settings/tokens")

# Initialize HF handler if key provided
hf_handler = None
if hf_api_key:
    hf_handler = HuggingFaceHandler(hf_api_key)
    st.sidebar.success("Hugging Face API connected!")

st.sidebar.markdown("---")

# Navigation
page = st.sidebar.radio(
    "Navigate",
    ["Dashboard", "Clients", "Events", "Venues", "Vendors", "Team", "AI Creative", "Proposal Generator"]
)


# Dashboard
if page == "Dashboard":
    st.markdown('<h1 class="main-header">EventOps Dashboard</h1>', unsafe_allow_html=True)
    st.markdown('<p class="hero-text">Run the show. Not six WhatsApp groups.</p>', unsafe_allow_html=True)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Clients", len(st.session_state.clients))
    with col2:
        st.metric("Events", len(st.session_state.events))
    with col3:
        st.metric("Venues", len(st.session_state.venues))
    with col4:
        st.metric("Vendors", len(st.session_state.vendors))
    with col5:
        st.metric("Team Members", len(st.session_state.team_members))
    
    st.markdown("---")
    st.subheader("Quick Actions")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Create New Event", use_container_width=True):
            st.session_state.selected_page = "Events"
            st.rerun()
    with col2:
        if st.button("Add Client", use_container_width=True):
            st.session_state.selected_page = "Clients"
            st.rerun()
    with col3:
        if st.button("Generate Proposal", use_container_width=True):
            st.session_state.selected_page = "Proposal Generator"
            st.rerun()
    
    st.markdown("---")
    st.subheader("Recent Events")
    if st.session_state.events:
        df_events = pd.DataFrame(st.session_state.events)
        if not df_events.empty:
            df_events = df_events.sort_values('created_at', ascending=False).head(5)
            st.dataframe(df_events[['name', 'event_type', 'event_date', 'status']], use_container_width=True)
    else:
        st.info("No events yet. Create your first event!")


# Clients
elif page == "Clients":
    st.markdown('<h1 class="main-header">Client Management</h1>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("All Clients")
        if st.session_state.clients:
            df_clients = pd.DataFrame(st.session_state.clients)
            st.dataframe(df_clients, use_container_width=True)
        else:
            st.info("No clients found")
    
    with col2:
        st.subheader("Add New Client")
        with st.form("add_client"):
            name = st.text_input("Name *")
            email = st.text_input("Email *")
            phone = st.text_input("Phone *")
            company = st.text_input("Company")
            address = st.text_area("Address")
            
            submitted = st.form_submit_button("Add Client")
            if submitted and name and email and phone:
                client_data = {
                    "name": name,
                    "email": email,
                    "phone": phone,
                    "company": company,
                    "address": address
                }
                created_client = excel_handler.create_client(client_data)
                st.session_state.clients = excel_handler.get_clients()
                st.success("Client added successfully!")
                st.rerun()


# Events with Sub-events (Wedding functions)
elif page == "Events":
    st.markdown('<h1 class="main-header">Event Management</h1>', unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["All Events", "Create Event", "Event Details"])
    
    with tab1:
        if st.session_state.events:
            df_events = pd.DataFrame(st.session_state.events)
            st.dataframe(df_events, use_container_width=True)
        else:
            st.info("No events found")
    
    with tab2:
        with st.form("create_event"):
            col1, col2 = st.columns(2)
            
            with col1:
                client_options = {c['name']: c['id'] for c in st.session_state.clients} if st.session_state.clients else {}
                client_name = st.selectbox("Client *", list(client_options.keys()))
                event_name = st.text_input("Event Name *")
                event_type = st.selectbox("Event Type", ["wedding", "corporate", "birthday", "conference", "other"])
                event_date = st.date_input("Event Date *")
                guest_count = st.number_input("Guest Count *", min_value=1)
            
            with col2:
                venue_options = {v['name']: v['id'] for v in st.session_state.venues} if st.session_state.venues else {}
                venue_name = st.selectbox("Venue", [""] + list(venue_options.keys()))
                status = st.selectbox("Status", ["enquiry", "planning", "proposed", "confirmed", "completed", "cancelled"])
                budget = st.number_input("Budget *", min_value=0.0, step=100.0)
                description = st.text_area("Description")
            
            submitted = st.form_submit_button("Create Event")
            if submitted:
                client_id = client_options.get(client_name)
                venue_id = venue_options.get(venue_name) if venue_name else None
                
                event_data = {
                    "client_id": client_id,
                    "name": event_name,
                    "event_type": event_type,
                    "event_date": event_date.isoformat(),
                    "guest_count": guest_count,
                    "venue_id": venue_id,
                    "status": status,
                    "budget": budget,
                    "description": description
                }
                
                created_event = excel_handler.create_event(event_data)
                st.session_state.events = excel_handler.get_events()
                st.success("Event created successfully!")
                st.rerun()
    
    with tab3:
        if st.session_state.events:
            event_options = {e['name']: e['id'] for e in st.session_state.events}
            selected_event = st.selectbox("Select Event", list(event_options.keys()))
            
            if selected_event:
                event_id = event_options[selected_event]
                
                # Sub-events for weddings (Mehendi, Sangeet, Wedding, Reception)
                st.subheader("Sub-Events (Functions)")
                sub_events = excel_handler.get_sub_events(event_id)
                if sub_events:
                    df_sub_events = pd.DataFrame(sub_events)
                    st.dataframe(df_sub_events, use_container_width=True)
                else:
                    st.info("No sub-events yet")
                
                with st.expander("Add Sub-Event"):
                    with st.form("add_sub_event"):
                        sub_event_name = st.text_input("Function Name *")
                        sub_event_date = st.date_input("Function Date *")
                        sub_event_type = st.selectbox("Function Type", ["Mehendi", "Sangeet", "Wedding", "Reception", "Other"])
                        submitted = st.form_submit_button("Add Function")
                        if submitted:
                            sub_event_data = {
                                "event_id": event_id,
                                "name": sub_event_name,
                                "date": sub_event_date.isoformat(),
                                "type": sub_event_type
                            }
                            excel_handler.create_sub_event(sub_event_data)
                            st.rerun()
                
                st.markdown("---")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.subheader("Budget")
                    budget_items = excel_handler.get_budget_items(event_id)
                    if budget_items:
                        df_budget = pd.DataFrame(budget_items)
                        st.dataframe(df_budget, use_container_width=True)
                    else:
                        st.info("No budget items")
                    
                    with st.expander("Add Budget Item"):
                        with st.form("add_budget_item"):
                            category = st.selectbox("Category", [
                                "Venue", "Catering", "Decoration", "Photography", 
                                "Entertainment", "Flex Printing", "Digital Printing",
                                "Truss", "Fabrication", "Genset", "Power", "Other"
                            ])
                            description = st.text_input("Description")
                            estimated_cost = st.number_input("Estimated Cost", min_value=0.0)
                            submitted = st.form_submit_button("Add")
                            if submitted:
                                item_data = {
                                    "event_id": event_id,
                                    "category": category,
                                    "description": description,
                                    "estimated_cost": estimated_cost
                                }
                                excel_handler.create_budget_item(item_data)
                                st.rerun()
                
                with col2:
                    st.subheader("Checklist")
                    tasks = excel_handler.get_tasks(event_id)
                    if tasks:
                        df_tasks = pd.DataFrame(tasks)
                        st.dataframe(df_tasks, use_container_width=True)
                    else:
                        st.info("No tasks")
                    
                    with st.expander("Add Task"):
                        with st.form("add_task"):
                            title = st.text_input("Task Title")
                            assigned_to = st.selectbox("Assigned To", [m['name'] for m in st.session_state.team_members] + ["Unassigned"])
                            due_date = st.date_input("Due Date")
                            priority = st.selectbox("Priority", ["low", "medium", "high", "urgent"])
                            submitted = st.form_submit_button("Add")
                            if submitted:
                                task_data = {
                                    "event_id": event_id,
                                    "title": title,
                                    "assigned_to": assigned_to,
                                    "due_date": due_date.isoformat(),
                                    "priority": priority
                                }
                                excel_handler.create_task(task_data)
                                st.rerun()
                
                with col3:
                    st.subheader("Runsheet")
                    runsheet = excel_handler.get_runsheet(event_id)
                    if runsheet:
                        df_runsheet = pd.DataFrame(runsheet)
                        st.dataframe(df_runsheet, use_container_width=True)
                    else:
                        st.info("No runsheet items")
                    
                    with st.expander("Add Runsheet Item"):
                        with st.form("add_runsheet"):
                            time = st.time_input("Time")
                            activity = st.text_input("Activity")
                            location = st.text_input("Location")
                            responsible = st.text_input("Responsible Person")
                            submitted = st.form_submit_button("Add")
                            if submitted:
                                item_data = {
                                    "event_id": event_id,
                                    "time": str(time),
                                    "activity": activity,
                                    "location": location,
                                    "responsible": responsible
                                }
                                excel_handler.create_runsheet_item(item_data)
                                st.rerun()


# Venues
elif page == "Venues":
    st.markdown('<h1 class="main-header">Venue Management</h1>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("All Venues")
        if st.session_state.venues:
            df_venues = pd.DataFrame(st.session_state.venues)
            st.dataframe(df_venues, use_container_width=True)
        else:
            st.info("No venues found")
    
    with col2:
        st.subheader("Add New Venue")
        with st.form("add_venue"):
            name = st.text_input("Venue Name *")
            address = st.text_input("Address *")
            capacity = st.number_input("Capacity *", min_value=1)
            contact_person = st.text_input("Contact Person *")
            contact_phone = st.text_input("Contact Phone *")
            email = st.text_input("Email *")
            hourly_rate = st.number_input("Hourly Rate *", min_value=0.0)
            confidence_score = st.slider("Confidence Score", 0.0, 1.0, 0.8, 0.1)
            notes = st.text_area("Notes")
            
            submitted = st.form_submit_button("Add Venue")
            if submitted:
                venue_data = {
                    "name": name,
                    "address": address,
                    "capacity": capacity,
                    "contact_person": contact_person,
                    "contact_phone": contact_phone,
                    "email": email,
                    "hourly_rate": hourly_rate,
                    "confidence_score": confidence_score,
                    "notes": notes
                }
                created_venue = excel_handler.create_venue(venue_data)
                st.session_state.venues = excel_handler.get_venues()
                st.success("Venue added successfully!")
                st.rerun()


# Vendors with specific categories
elif page == "Vendors":
    st.markdown('<h1 class="main-header">Vendor Management</h1>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("All Vendors")
        if st.session_state.vendors:
            df_vendors = pd.DataFrame(st.session_state.vendors)
            st.dataframe(df_vendors, use_container_width=True)
        else:
            st.info("No vendors found")
    
    with col2:
        st.subheader("Add New Vendor")
        with st.form("add_vendor"):
            name = st.text_input("Vendor Name *")
            category = st.selectbox("Category", [
                "Catering", "Photography", "Decoration", "Entertainment", 
                "Audio Visual", "Flex Printing", "Digital Printing",
                "Truss", "Fabrication", "Genset", "Power", "Lighting", "Sound", "Other"
            ])
            contact_person = st.text_input("Contact Person *")
            contact_phone = st.text_input("Contact Phone *")
            email = st.text_input("Email *")
            rate = st.number_input("Rate *", min_value=0.0)
            rating = st.slider("Rating", 0.0, 5.0, 0.0, 0.5)
            notes = st.text_area("Notes")
            
            submitted = st.form_submit_button("Add Vendor")
            if submitted:
                vendor_data = {
                    "name": name,
                    "category": category,
                    "contact_person": contact_person,
                    "contact_phone": contact_phone,
                    "email": email,
                    "rate": rate,
                    "rating": rating if rating > 0 else None,
                    "notes": notes
                }
                created_vendor = excel_handler.create_vendor(vendor_data)
                st.session_state.vendors = excel_handler.get_vendors()
                st.success("Vendor added successfully!")
                st.rerun()


# Team Management
elif page == "Team":
    st.markdown('<h1 class="main-header">Team Management</h1>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Team Members")
        if st.session_state.team_members:
            df_team = pd.DataFrame(st.session_state.team_members)
            st.dataframe(df_team, use_container_width=True)
        else:
            st.info("No team members found")
    
    with col2:
        st.subheader("Add Team Member")
        with st.form("add_team_member"):
            name = st.text_input("Name *")
            role = st.selectbox("Role", [
                "Event Manager", "Creative Director", "Production Manager",
                "Coordinator", "Designer", "Account Manager", "Other"
            ])
            email = st.text_input("Email *")
            phone = st.text_input("Phone")
            
            submitted = st.form_submit_button("Add Team Member")
            if submitted and name and email:
                member_data = {
                    "name": name,
                    "role": role,
                    "email": email,
                    "phone": phone
                }
                created_member = excel_handler.create_team_member(member_data)
                st.session_state.team_members = excel_handler.get_team_members()
                st.success("Team member added successfully!")
                st.rerun()


# AI Creative (Logo, Stage Design, Moodboard)
elif page == "AI Creative":
    st.markdown('<h1 class="main-header">AI Creative Studio</h1>', unsafe_allow_html=True)
    st.markdown('<p class="hero-text">Generate coherent creative direction with AI</p>', unsafe_allow_html=True)
    
    if not hf_handler:
        st.warning("Please enter your Hugging Face API key in the sidebar to use AI features")
    else:
        creative_type = st.selectbox(
            "Creative Type",
            ["Stage Design", "Logo Design", "Moodboard", "Color Palette", "Lighting Design"]
        )
        
        event_options = {e['name']: e for e in st.session_state.events} if st.session_state.events else {}
        selected_event = st.selectbox("Select Event (Optional)", ["None"] + list(event_options.keys()))
        
        if selected_event != "None":
            context = event_options[selected_event]
            st.json(context)
        
        prompt = st.text_area(
            "Describe your creative vision",
            placeholder="E.g., A Bollywood-themed conference stage with gold and red colors, modern truss structure...",
            height=150
        )
        
        if st.button("Generate Creative Direction"):
            if prompt:
                with st.spinner("AI is generating creative direction..."):
                    # Use different models based on creative type
                    models = {
                        "Stage Design": "mistralai/Mistral-7B-Instruct-v0.2",
                        "Logo Design": "mistralai/Mistral-7B-Instruct-v0.2",
                        "Moodboard": "mistralai/Mistral-7B-Instruct-v0.2",
                        "Color Palette": "mistralai/Mistral-7B-Instruct-v0.2",
                        "Lighting Design": "mistralai/Mistral-7B-Instruct-v0.2"
                    }
                    
                    model = models.get(creative_type, "mistralai/Mistral-7B-Instruct-v0.2")
                    
                    system_prompt = f"""
                    You are an expert event creative designer specializing in {creative_type}.
                    Generate detailed, specific creative direction based on the user's description.
                    Include technical details like truss layout, lighting colors, floral arrangements, typography, etc.
                    Be specific and practical for event production.
                    """
                    
                    full_prompt = f"{system_prompt}\n\nUser request: {prompt}"
                    
                    result = hf_handler.generate_text(model, full_prompt, max_length=800)
                    
                    st.subheader("AI-Generated Creative Direction")
                    st.write(result)
                    
                    if selected_event != "None":
                        if st.button("Save to Event"):
                            direction_data = {
                                "event_id": event_options[selected_event]['id'],
                                "type": creative_type,
                                "prompt": prompt,
                                "result": result
                            }
                            excel_handler.create_creative_direction(direction_data)
                            st.success("Creative direction saved!")
            else:
                st.warning("Please enter a description")


# Proposal Generator
elif page == "Proposal Generator":
    st.markdown('<h1 class="main-header">Proposal Generator</h1>', unsafe_allow_html=True)
    st.markdown('<p class="hero-text">Build a proposal your client can actually picture — in minutes</p>', unsafe_allow_html=True)
    
    if not hf_handler:
        st.warning("Please enter your Hugging Face API key in the sidebar to use AI features")
    else:
        event_options = {e['name']: e for e in st.session_state.events} if st.session_state.events else {}
        selected_event = st.selectbox("Select Event", list(event_options.keys()))
        
        if selected_event:
            event_id = event_options[selected_event]['id']
            event_data = event_options[selected_event]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Proposal Details")
                proposal_title = st.text_input("Proposal Title", value=f"Proposal for {selected_event}")
                client_name = st.text_input("Client Name", value=event_data.get('client_id', ''))
            
            with col2:
                st.subheader("Proposal Sections")
                include_moodboard = st.checkbox("Include Moodboard", value=True)
                include_creative = st.checkbox("Include Creative Direction", value=True)
                include_budget = st.checkbox("Include Cost Breakdown", value=True)
                include_timeline = st.checkbox("Include Timeline", value=True)
            
            additional_notes = st.text_area("Additional Notes for AI", placeholder="Any specific requirements or themes...")
            
            if st.button("Generate Proposal"):
                with st.spinner("AI is generating your proposal..."):
                    prompt = f"""
                    Generate a professional event proposal with the following details:
                    Event: {selected_event}
                    Type: {event_data.get('event_type')}
                    Date: {event_data.get('event_date')}
                    Guest Count: {event_data.get('guest_count')}
                    Budget: {event_data.get('budget')}
                    
                    Include sections:
                    - Executive Summary
                    - Event Overview
                    - Creative Direction
                    - Moodboard Description
                    - Cost Breakdown
                    - Timeline
                    - Terms and Conditions
                    
                    Additional Notes: {additional_notes}
                    
                    Make it professional, persuasive, and detailed.
                    """
                    
                    result = hf_handler.generate_text("mistralai/Mistral-7B-Instruct-v0.2", prompt, max_length=1500)
                    
                    st.subheader("Generated Proposal")
                    st.write(result)
                    
                    st.download_button(
                        "Download Proposal",
                        result,
                        file_name=f"{proposal_title}.txt",
                        mime="text/plain"
                    )


# Footer
st.sidebar.markdown("---")
st.sidebar.markdown("© 2024 EventOps")
