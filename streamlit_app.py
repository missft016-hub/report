import streamlit as st
import requests
import pandas as pd
from datetime import datetime, date
import json
from bs4 import BeautifulSoup
import re
import os
import io
import platform
from typing import Optional, Dict, Any

# Set page config
st.set_page_config(
    page_title="Teams Message Fetcher",
    page_icon="💬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====== UTILITY FUNCTIONS ======
def clean_teams_message(html_content: str) -> str:
    """Clean HTML content from Teams messages and convert to plain text."""
    soup = BeautifulSoup(html_content or "", "html.parser")
    text = soup.get_text(separator=" ", strip=True)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'<at>(.*?)</at>', r'@\1', text)
    text = text.encode("ascii", errors="ignore").decode()
    return text.strip()

def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Get Microsoft Graph API access token."""
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    
    try:
        response = requests.post(url, data=data)
        if response.status_code != 200:
            raise Exception(f"Failed to get token: {response.status_code}, {response.text}")
        return response.json()["access_token"]
    except Exception as e:
        st.error(f"Authentication failed: {str(e)}")
        return None

def get_graph_headers(token: str) -> Dict[str, str]:
    """Get headers for Microsoft Graph API requests."""
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

def fetch_replies_done(headers: Dict[str, str], team_id: str, channel_id: str, message_id: str) -> Optional[Dict[str, Any]]:
    """Fetch replies to check for 'done' messages."""
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
    done_pattern = re.compile(r'\bdone\b', re.IGNORECASE)

    try:
        while url:
            resp = requests.get(url, headers=headers)
            if resp.status_code == 404:
                return None
            if resp.status_code != 200:
                return None
            
            data = resp.json()
            replies = data.get("value", [])
            
            if not replies:
                return None

            for reply in replies:
                content = reply.get("body", {}).get("content", "") or ""
                if done_pattern.search(content):
                    user = reply.get("from", {}).get("user", {}).get("displayName", "Unknown")
                    created = reply.get("createdDateTime", "")
                    return {
                        "user": user,
                        "content": content,
                        "createdDateTime": created
                    }

            url = data.get('@odata.nextLink')
        return None
    except Exception:
        return None

def fetch_channel_messages(headers: Dict[str, str], team_id: str, channel_id: str, progress_bar=None) -> pd.DataFrame:
    """Fetch messages from Teams channel and return as DataFrame."""
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages"
    rows = []
    page_count = 0
    
    while url:
        try:
            resp = requests.get(url, headers=headers)
            if resp.status_code != 200:
                st.error(f"Graph API error: {resp.status_code}, {resp.text}")
                break

            data = resp.json()
            messages = data.get("value", [])
            
            for msg in messages:
                # Safe sender extraction
                from_user = msg.get("from")
                if from_user and "user" in from_user:
                    sender = from_user["user"].get("displayName", "Unknown User")
                elif from_user and "application" in from_user:
                    sender = from_user["application"].get("displayName", "App Message")
                else:
                    sender = "System or Unknown"

                # Get replies to check for task completion
                replies_done = fetch_replies_done(headers, team_id, channel_id, msg.get("id", ""))
                task_done = replies_done.get("createdDateTime") if replies_done else None

                # Message content
                raw_content = msg.get("body", {}).get("content", "")
                clean_message = clean_teams_message(raw_content)

                # Timestamp and subject/topic
                created = msg.get("createdDateTime", "")
                subject = msg.get("subject", "")

                # Enhanced type detection
                msg_type = "Request"  # Default type
                subject_lower = subject.lower() if subject else ""
                
                if re.search(r'\b(bug|error|fail|broken)\b', subject_lower):
                    msg_type = "Bug"
                elif re.search(r'\b(issue|problem|trouble)\b', subject_lower):
                    msg_type = "Issue"
                elif re.search(r'\b(request|support|help|fw|forward|fwd)\b', subject_lower):
                    msg_type = "Request"

                # Attachments
                attachments = msg.get("attachments") or []
                attachment_names = []
                attachment_urls = []

                for a in attachments:
                    name = a.get("name", "unnamed")
                    content_url = a.get("contentUrl")
                    if content_url:
                        attachment_names.append(name)
                        attachment_urls.append(content_url)

                # Calculate task duration
                estimated_duration = "0 days, 10 hours, 0 minutes"
                estimated_hours = 10.0
                
                if task_done is not None:
                    try:
                        created_time = pd.to_datetime(created).tz_localize(None)
                        done_time = pd.to_datetime(task_done).tz_localize(None)
                        duration_timedelta = done_time - created_time
                        
                        days = max(0, duration_timedelta.days)
                        seconds = max(0, duration_timedelta.seconds)
                        hours = seconds // 3600
                        minutes = (seconds % 3600) // 60
                        
                        if days == 0 and hours == 0 and minutes == 0:
                            estimated_duration = "0 days, 10 hours, 0 minutes"
                            estimated_hours = 10.0
                        else:
                            estimated_duration = f"{days} days, {hours} hours, {minutes} minutes"
                            estimated_hours = round(max(0, duration_timedelta.total_seconds() / 3600), 2)
                    except Exception:
                        estimated_duration = "0 days, 10 hours, 0 minutes"
                        estimated_hours = 10.0

                # Add row
                rows.append({
                    "Sender": sender,
                    "Topic": subject,
                    "Message": clean_message,
                    "Date": created,
                    "Type": msg_type,
                    "AttachmentNames": ", ".join(attachment_names),
                    "AttachmentURLs": "\n".join(attachment_urls),
                    "DoneDate": task_done,
                    "EstimatedHours": estimated_hours,
                    "EstimatedDuration": estimated_duration
                })

            url = data.get("@odata.nextLink")
            page_count += 1
            
            if progress_bar:
                progress_bar.progress(min(page_count * 10, 90))  # Approximate progress
                
        except Exception as e:
            st.error(f"Error fetching messages: {str(e)}")
            break

    return pd.DataFrame(rows)

def create_summary_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Create summary statistics from the messages DataFrame."""
    if df.empty:
        return pd.DataFrame()
    
    rows = []
    total_messages = len(df)
    
    # Summary by type
    if "Type" in df.columns:
        unique_types = df["Type"].unique()
        for msg_type in unique_types:
            type_df = df[df["Type"] == msg_type]
            count = len(type_df)
            avg_hours = type_df["EstimatedHours"].mean() if "EstimatedHours" in type_df.columns else 0
            
            rows.append({
                "Type": msg_type,
                "Total": count,
                "AvgEstimatedHours": round(avg_hours, 2),
                "Percentage": round((count / total_messages) * 100, 1)
            })
    
    # Total row
    total_avg_hours = df["EstimatedHours"].mean() if "EstimatedHours" in df.columns else 0
    rows.append({
        "Type": "Total",
        "Total": total_messages,
        "AvgEstimatedHours": round(total_avg_hours, 2),
        "Percentage": 100.0
    })
    
    return pd.DataFrame(rows)

def filter_dataframe(df: pd.DataFrame, date_from: date = None, date_to: date = None,
                     sender_filter: str = "", show_all: bool = False) -> pd.DataFrame:
    """Filter DataFrame based on date range and sender."""
    if df.empty:
        return df
    
    try:
        # Convert date column
        df["Date"] = pd.to_datetime(df["Date"]).dt.tz_localize(None)
        filtered_df = df.copy()
        
        # Apply date filters if not showing all messages
        if not show_all and date_from and date_to:
            filter_date_from = pd.to_datetime(date_from)
            filter_date_to = pd.to_datetime(date_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)  # End of day
            
            filtered_df = filtered_df[
                (filtered_df["Date"] >= filter_date_from) &
                (filtered_df["Date"] <= filter_date_to)
            ]
        elif not show_all and date_from:
            filter_date_from = pd.to_datetime(date_from)
            filtered_df = filtered_df[filtered_df["Date"] >= filter_date_from]
        
        # Apply sender filter
        if sender_filter:
            filtered_df = filtered_df[
                filtered_df["Sender"].str.contains(sender_filter, case=False, na=False)
            ]
            
        return filtered_df
    except Exception as e:
        st.error(f"Error filtering data: {str(e)}")
        return df

def export_to_excel(df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    """Export DataFrames to Excel and return as bytes."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Messages', index=False)
        if not summary_df.empty:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    output.seek(0)
    return output.getvalue()

def post_to_jira(df: pd.DataFrame, jira_url: str, jira_email: str, jira_token: str,
                 project_key: str, epic_key: str, progress_callback=None) -> Dict[str, Any]:
    """Post messages to Jira as tasks under an epic."""
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    auth = (jira_email, jira_token)
    
    results = {
        "success": 0,
        "failed": 0,
        "errors": []
    }
    
    total_messages = len(df)
    
    for idx, (_, row) in enumerate(df.iterrows()):
        try:
            # Format date
            try:
                date_obj = pd.to_datetime(row['Date'])
                formatted_date = date_obj.strftime("%Y-%m-%d %H:%M")
            except Exception:
                formatted_date = str(row['Date'])

            # Build issue fields
            summary = f"{row['Topic']} - Teams message by {row['Sender']} at {formatted_date}"
            plain_text = str(row['Message'])

            # Convert to Atlassian Document Format (ADF)
            description_adf = {
                "type": "doc",
                "version": 1,
                "content": [
                    {
                        "type": "paragraph",
                        "content": [
                            {
                                "type": "text",
                                "text": plain_text
                            }
                        ]
                    }
                ]
            }

            payload = {
                "fields": {
                    "project": {"key": project_key},
                    "summary": summary,
                    "description": description_adf,
                    "issuetype": {"id": "10089"},  # Task type ID
                    "parent": {"key": epic_key}
                }
            }

            # Send request
            response = requests.post(
                f"{jira_url}/rest/api/3/issue",
                headers=headers,
                auth=auth,
                data=json.dumps(payload)
            )

            if response.status_code == 201:
                issue_key = response.json().get("key")
                results["success"] += 1
                if progress_callback:
                    progress_callback(idx + 1, total_messages, f"Created {issue_key}")
            else:
                results["failed"] += 1
                error_msg = f"Row {idx + 1}: {response.status_code} - {response.text[:100]}..."
                results["errors"].append(error_msg)
                if progress_callback:
                    progress_callback(idx + 1, total_messages, f"Failed: {response.status_code}")

        except Exception as e:
            results["failed"] += 1
            error_msg = f"Row {idx + 1}: {str(e)}"
            results["errors"].append(error_msg)
            if progress_callback:
                progress_callback(idx + 1, total_messages, f"Error: {str(e)[:50]}...")
    
    return results

# ====== STREAMLIT APP ======
def main():
    st.title("💬 Teams Message Fetcher")
    st.markdown("Fetch and analyze Microsoft Teams channel messages with ease!")
    
    # Sidebar configuration
    st.sidebar.header("Configuration")
    
    with st.sidebar.expander("🔐 Authentication", expanded=True):
        tenant_id = st.text_input("Tenant ID", value="5b97cd8a-cbc2-4968-9f16-12ca0fd57013")
        client_id = st.text_input("Client ID", value="33d58b6e-49f1-4c50-9f83-6b8b28d22fdb")
        client_secret = st.text_input("Client Secret", type="password", value="r2u8Q~2HVZNHGRC3~MPpP83BdMMbrbUYdF1-mbvt")
    
    with st.sidebar.expander("📊 Teams Configuration", expanded=True):
        team_id = st.text_input("Team ID", value="22e6688f-a1d9-4d15-9103-a1ec37194022")
        channel_id = st.text_input("Channel ID", value="19:27f5773cb5f74f5488e03320cbcf447c@thread.tacv2")
    
    with st.sidebar.expander("🔍 Filters", expanded=True):
        # Date range filter
        col1, col2 = st.columns(2)
        with col1:
            date_from = st.date_input("From date", value=date(2024, 1, 1))
        with col2:
            date_to = st.date_input("To date", value=date.today())
        
        # Sender filter
        sender_filter = st.text_input("Filter by sender (optional)", placeholder="e.g., Ali Murtadho")
        
        # Show all messages option
        show_all_messages = st.checkbox("Show all messages (ignore date filter)", value=False)
    
    with st.sidebar.expander("🎫 Jira Integration (Optional)", expanded=False):
        enable_jira = st.checkbox("Enable Jira Integration")
        if enable_jira:
            jira_url = st.text_input("Jira URL", value="https://lnk-team-ixag51a4.atlassian.net")
            jira_email = st.text_input("Jira Email", value="mis.freelance@lnk.co.id")
            jira_token = st.text_input("Jira API Token", type="password",
                                     value="ATATT3xFfGF0OaK_RD5uZWt1ABiE4QBK5LufO1g4BXytYajoISJknAQ13boixEeTUqC2Ex_XpYJbMeGoOyGQNVx9PAKF5kmi3CgLWN3FJj0ZSeo3Wm0z1e3FEHx6cTdbiKi-inyVw09QbVtrtV61pzpsI2A1b25D2Amj4SfSDraWg8CViVQnxJM=6EF6C3DB")
            jira_project_key = st.text_input("Project Key", value="SPONELTL")
            jira_epic_key = st.text_input("Epic Key", value="SPONELTL-9")
        else:
            jira_url = jira_email = jira_token = jira_project_key = jira_epic_key = ""
    
    # Main content
    col1, col2 = st.columns([3, 1])
    
    with col2:
        fetch_button = st.button("🚀 Fetch Messages", type="primary", use_container_width=True)
    
    with col1:
        st.markdown("Click **Fetch Messages** to retrieve and analyze Teams channel messages")
    
    # Session state for storing data
    if "messages_df" not in st.session_state:
        st.session_state.messages_df = pd.DataFrame()
    if "summary_df" not in st.session_state:
        st.session_state.summary_df = pd.DataFrame()
    
    # Fetch messages
    if fetch_button:
        if not all([tenant_id, client_id, client_secret, team_id, channel_id]):
            st.error("Please fill in all required configuration fields.")
            return
        
        with st.spinner("Authenticating with Microsoft Graph API..."):
            token = get_access_token(tenant_id, client_id, client_secret)
        
        if token:
            st.success("✅ Authentication successful!")
            headers = get_graph_headers(token)
            
            with st.spinner("Fetching messages from Teams channel..."):
                progress_bar = st.progress(0)
                messages_df = fetch_channel_messages(headers, team_id, channel_id, progress_bar)
                progress_bar.progress(100)
            
            if not messages_df.empty:
                st.success(f"✅ Fetched {len(messages_df)} messages successfully!")
                
                # Apply filters
                filtered_df = filter_dataframe(messages_df, date_from, date_to, sender_filter, show_all_messages)
                summary_df = create_summary_dataframe(filtered_df)
                
                # Store in session state
                st.session_state.messages_df = filtered_df
                st.session_state.summary_df = summary_df
                
                # Show filter status
                if show_all_messages:
                    filter_info = f"📊 Showing all {len(filtered_df)} messages (date filter disabled)"
                else:
                    date_range = f"{date_from.strftime('%Y-%m-%d')} to {date_to.strftime('%Y-%m-%d')}"
                    filter_info = f"📊 Showing {len(filtered_df)} messages from {date_range}"
                    if sender_filter:
                        filter_info += f" | Sender: '{sender_filter}'"
                
                if len(filtered_df) == 0:
                    st.warning("⚠️ No messages found with current filters. Try:")
                    st.markdown("- Expanding the date range")
                    st.markdown("- Checking 'Show all messages' option")
                    st.markdown("- Removing sender filter")
                else:
                    st.info(filter_info)
            else:
                st.warning("No messages found or error occurred while fetching.")
    
    # Display results
    if not st.session_state.messages_df.empty:
        st.markdown("---")
        
        # Summary section
        st.header("📈 Summary")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Messages", len(st.session_state.messages_df))
        with col2:
            avg_hours = st.session_state.messages_df["EstimatedHours"].mean()
            st.metric("Avg Hours", f"{avg_hours:.1f}")
        with col3:
            total_hours = st.session_state.messages_df["EstimatedHours"].sum()
            st.metric("Total Hours", f"{total_hours:.1f}")
        with col4:
            unique_senders = st.session_state.messages_df["Sender"].nunique()
            st.metric("Unique Senders", unique_senders)
        
        # Summary table
        if not st.session_state.summary_df.empty:
            st.subheader("Summary by Type")
            st.dataframe(st.session_state.summary_df, use_container_width=True)
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            if "Type" in st.session_state.messages_df.columns:
                st.subheader("Messages by Type")
                type_counts = st.session_state.messages_df["Type"].value_counts()
                st.bar_chart(type_counts)
        
        with col2:
            if "Date" in st.session_state.messages_df.columns:
                st.subheader("Messages Over Time")
                daily_counts = st.session_state.messages_df.groupby(
                    st.session_state.messages_df["Date"].dt.date
                ).size()
                st.line_chart(daily_counts)
        
        # Messages table
        st.subheader("📋 Messages")
        
        # Additional filters for the table
        col1, col2, col3 = st.columns(3)
        with col1:
            selected_types = st.multiselect(
                "Filter by Type",
                options=st.session_state.messages_df["Type"].unique(),
                default=st.session_state.messages_df["Type"].unique()
            )
        with col2:
            selected_senders = st.multiselect(
                "Filter by Sender",
                options=st.session_state.messages_df["Sender"].unique(),
                default=st.session_state.messages_df["Sender"].unique()
            )
        with col3:
            max_rows = st.selectbox("Max rows to display", [50, 100, 200, 500], index=1)
        
        # Apply additional filters
        display_df = st.session_state.messages_df[
            (st.session_state.messages_df["Type"].isin(selected_types)) &
            (st.session_state.messages_df["Sender"].isin(selected_senders))
        ].head(max_rows)
        
        # Display table
        st.dataframe(
            display_df,
            use_container_width=True,
            column_config={
                "Date": st.column_config.DatetimeColumn("Date", format="DD/MM/YYYY HH:mm"),
                "EstimatedHours": st.column_config.NumberColumn("Hours", format="%.1f"),
                "Message": st.column_config.TextColumn("Message", width="large")
            }
        )
        
        # Export section
        st.markdown("---")
        st.subheader("📥 Export & Integration")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("📊 Download Excel", use_container_width=True):
                excel_data = export_to_excel(st.session_state.messages_df, st.session_state.summary_df)
                st.download_button(
                    label="💾 Download teams_messages.xlsx",
                    data=excel_data,
                    file_name=f"teams_messages_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with col2:
            csv_data = st.session_state.messages_df.to_csv(index=False)
            st.download_button(
                label="📄 Download CSV",
                data=csv_data,
                file_name=f"teams_messages_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col3:
            # Jira Integration
            if enable_jira and all([jira_url, jira_email, jira_token, jira_project_key, jira_epic_key]):
                if st.button("🎫 Post to Jira", use_container_width=True):
                    with st.spinner("Posting messages to Jira..."):
                        progress_container = st.container()
                        status_container = st.container()
                        
                        def progress_callback(current, total, message):
                            progress_pct = int((current / total) * 100)
                            progress_container.progress(progress_pct)
                            status_container.info(f"Progress: {current}/{total} - {message}")
                        
                        try:
                            results = post_to_jira(
                                st.session_state.messages_df,
                                jira_url, jira_email, jira_token,
                                jira_project_key, jira_epic_key,
                                progress_callback
                            )
                            
                            # Show results
                            if results["success"] > 0:
                                st.success(f"✅ Successfully created {results['success']} Jira tasks!")
                            
                            if results["failed"] > 0:
                                st.error(f"❌ Failed to create {results['failed']} tasks")
                                
                                # Show errors in expander
                                with st.expander("View Errors", expanded=False):
                                    for error in results["errors"][:10]:  # Show first 10 errors
                                        st.text(error)
                                    if len(results["errors"]) > 10:
                                        st.text(f"... and {len(results['errors']) - 10} more errors")
                            
                        except Exception as e:
                            st.error(f"Error posting to Jira: {str(e)}")
            elif enable_jira:
                st.warning("⚠️ Please fill in all Jira configuration fields to enable posting")

if __name__ == "__main__":
    main()