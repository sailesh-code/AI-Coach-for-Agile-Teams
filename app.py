from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from jira import JIRA
import os
from dotenv import load_dotenv
import google.generativeai as genai
from datetime import datetime, timedelta
import json
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import pandas as pd
import numpy as np

# Load environment variables
load_dotenv()

app = Flask(__name__)
CORS(app)

# Configure Gemini
genai.configure(api_key=os.getenv('GEMINI_API_KEY'))
model = genai.GenerativeModel('gemini-2.0-flash')

# Jira configuration
JIRA_URL = os.getenv('JIRA_URL')
JIRA_EMAIL = os.getenv('JIRA_EMAIL')
JIRA_API_TOKEN = os.getenv('JIRA_API_TOKEN')

def get_jira_client():
    if not all([JIRA_URL, JIRA_EMAIL, JIRA_API_TOKEN]):
        raise ValueError("Missing Jira configuration. Please check your .env file.")
    
    # Ensure the URL doesn't end with a slash
    jira_url = JIRA_URL.rstrip('/')
    
    try:
        return JIRA(
            server=jira_url,
            basic_auth=(JIRA_EMAIL, JIRA_API_TOKEN),
            validate=True
        )
    except Exception as e:
        raise Exception(f"Failed to connect to Jira: {str(e)}")

@app.route('/api/boards', methods=['GET'])
def get_boards():
    try:
        jira_client = get_jira_client()
        boards = jira_client.boards()
        return jsonify([{
            'id': board.id,
            'name': board.name,
            'type': board.type
        } for board in boards])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sprints', methods=['GET'])
def get_sprints():
    try:
        board_id = request.args.get('boardId')
        if not board_id:
            return jsonify({'error': 'Board ID is required'}), 400

        jira_client = get_jira_client()
        sprints = jira_client.sprints(board_id)
        
        # Format sprint data
        formatted_sprints = []
        for sprint in sprints:
            formatted_sprints.append({
                'id': sprint.id,
                'name': sprint.name,
                'state': sprint.state,
                'startDate': sprint.startDate,
                'endDate': sprint.endDate,
                'goal': sprint.goal if hasattr(sprint, 'goal') else None
            })
        
        # Sort sprints by end date (most recent first)
        formatted_sprints.sort(key=lambda x: x['endDate'] if x['endDate'] else datetime.min, reverse=True)
        
        return jsonify(formatted_sprints)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def get_sprint_stories(jira_client, sprint_id):
    # JQL query to get all stories in the sprint
    jql = f'sprint = {sprint_id} AND type in (Story, Task, Bug) ORDER BY created DESC'
    issues = jira_client.search_issues(jql, maxResults=False, expand='changelog,renderedFields')
    
    stories = []
    for issue in issues:
        # Get all available fields
        story_data = {
            'key': issue.key,
            'summary': issue.fields.summary,
            'description': issue.fields.description,
            'status': issue.fields.status.name,
            'type': issue.fields.issuetype.name,
            'priority': getattr(issue.fields, 'priority', None).name if getattr(issue.fields, 'priority', None) else None,
            'assignee': getattr(issue.fields.assignee, 'displayName', None) if issue.fields.assignee else None,
            'reporter': getattr(issue.fields.reporter, 'displayName', None) if issue.fields.reporter else None,
            'created': issue.fields.created,
            'updated': issue.fields.updated,
            'resolution': getattr(issue.fields.resolution, 'name', None) if issue.fields.resolution else None,
            'labels': getattr(issue.fields, 'labels', []),
            'components': [comp.name for comp in getattr(issue.fields, 'components', [])],
            'story_points': getattr(issue.fields, 'customfield_10016', None),  # Adjust field ID based on your Jira setup
            'epic_link': getattr(issue.fields, 'customfield_10014', None),  # Adjust field ID based on your Jira setup
            'subtasks': [],
            'changelog': [],
            'comments': [],
            'blockers': []
        }
        
        # Get subtasks
        subtasks = jira_client.search_issues(f'parent = {issue.key}', expand='changelog')
        for subtask in subtasks:
            subtask_data = {
                'key': subtask.key,
                'summary': subtask.fields.summary,
                'description': subtask.fields.description,
                'status': subtask.fields.status.name,
                'assignee': getattr(subtask.fields.assignee, 'displayName', None) if subtask.fields.assignee else None,
                'created': subtask.fields.created,
                'updated': subtask.fields.updated,
                'changelog': [],
                'blockers': []
            }
            
            # Get subtask changelog
            for history in subtask.changelog.histories:
                for item in history.items:
                    subtask_data['changelog'].append({
                        'date': history.created,
                        'author': history.author.displayName,
                        'field': item.field,
                        'from': item.fromString,
                        'to': item.toString
                    })
            
            # Get subtask blockers using issue links
            try:
                blocker_links = jira_client.search_issues(f'issue in linkedIssues({subtask.key}) AND type in (Bug, Story, Task)')
                for blocker in blocker_links:
                    subtask_data['blockers'].append({
                        'key': blocker.key,
                        'summary': blocker.fields.summary,
                        'status': blocker.fields.status.name
                    })
            except Exception as e:
                print(f"Error fetching blockers for subtask {subtask.key}: {str(e)}")
            
            story_data['subtasks'].append(subtask_data)
        
        # Get story changelog
        for history in issue.changelog.histories:
            for item in history.items:
                story_data['changelog'].append({
                    'date': history.created,
                    'author': history.author.displayName,
                    'field': item.field,
                    'from': item.fromString,
                    'to': item.toString
                })
        
        # Get comments
        if hasattr(issue.fields, 'comment'):
            for comment in issue.fields.comment.comments:
                story_data['comments'].append({
                    'author': comment.author.displayName,
                    'body': comment.body,
                    'created': comment.created
                })
        
        # Get story blockers using issue links
        try:
            blocker_links = jira_client.search_issues(f'issue in linkedIssues({issue.key}) AND type in (Bug, Story, Task)')
            for blocker in blocker_links:
                story_data['blockers'].append({
                    'key': blocker.key,
                    'summary': blocker.fields.summary,
                    'status': blocker.fields.status.name
                })
        except Exception as e:
            print(f"Error fetching blockers for story {issue.key}: {str(e)}")
        
        stories.append(story_data)
    
    return stories

def generate_subgoals(sprint_goal):
    prompt = f"""
    Do not summarize, rewrite, or rephrase any part of the text. Each subgoal should be exactly as it appears in the original sprint goal, just separated out clearly. Do not make up any subgoals. Do not split any sentences.
    Sprint Goal: {sprint_goal}
    """
    
    response = model.generate_content(prompt)
    return response.text

def assign_stories_to_subgoals(stories, subgoals):
    # Create a detailed prompt for the AI to assign stories to subgoals
    stories_text = "\n".join([
        f"Story {story['key']}:\n"
        f"Summary: {story['summary']}\n"
        f"Description: {story['description']}\n"
        f"Type: {story['type']}\n"
        f"Status: {story['status']}\n"
        f"Labels: {', '.join(story['labels'])}\n"
        f"Components: {', '.join(story['components'])}\n"
        f"Priority: {story['priority']}\n"
        for story in stories
    ])
    
    prompt = f"""
    You are a Product Owner analyzing stories from Jira. Below is the list of user stories and sprint goals for the current sprint.

    Your task is to assign each story to the most relevant sprint goal based on the story's description and summary and acceptance criteria. If a story could relate to multiple goals, assign it to the most appropriate primary goal. If a story does not relate to any of the goals, mark it as "Unassigned".
    
    Subgoals:
    {subgoals}
    
    Stories:
    {stories_text}
    
    Return the assignments in this format:
    Subgoal 1:
    - STORY-123: Story Summary
    - STORY-456: Story Summary

    Subgoal 2:
    - STORY-789: Story Summary
    
    Unassigned:
    - STORY-ABC: Story Summary
    """
    
    response = model.generate_content(prompt)
    return response.text

def generate_achievements(stories, subgoals):
    # Create a detailed prompt for analyzing stories and generating achievements
    stories_text = "\n".join([
        f"Story {story['key']}:\n"
        f"Summary: {story['summary']}\n"
        f"Description: {story['description']}\n"
        f"Type: {story['type']}\n"
        f"Status: {story['status']}\n"
        f"Labels: {', '.join(story['labels'])}\n"
        f"Components: {', '.join(story['components'])}\n"
        f"Priority: {story['priority']}\n"
        f"Comments: {len(story['comments'])} comments\n"
        for story in stories
    ])
    
    prompt = f"""
    Analyze the following stories and their assignments to subgoals. For each subgoal, identify key achievements.
    Focus on completed tasks, technical achievements, improvements, and measurable results.
    
    Subgoals:
    {subgoals}
    
    Stories:
    {stories_text}
    
    For each subgoal, list key achievements as bullet points.
    For each subgoal, list all the assigned stories as story numbers from subgoals variable.
    Format exactly like this:
    Subgoal 1:
    Story Numbers: STORY-123, STORY-456
    - First achievement
    - Second achievement
    - Third achievement
    
    Subgoal 2:
    Story Numbers: STORY-789
    - First achievement
    - Second achievement
    - Third achievement
    
    Important: Each achievement should be a complete sentence starting with a capital letter.
    Do not number the achievements or add any additional formatting.
    """
    
    response = model.generate_content(prompt)
    return response.text

@app.route('/api/sprint-report', methods=['GET'])
def get_sprint_report():
    try:
        board_id = request.args.get('boardId')
        sprint_id = request.args.get('sprintId')
        
        if not board_id or not sprint_id:
            return jsonify({'error': 'Board ID and Sprint ID are required'}), 400

        jira_client = get_jira_client()
        
        # Get sprint details
        sprint = jira_client.sprint(sprint_id)
        if not sprint:
            return jsonify({'error': 'Sprint not found'}), 404
        
        # Get sprint goal
        sprint_goal = sprint.goal if hasattr(sprint, 'goal') else "No sprint goal found"
        
        # Generate subgoals using Gemini
        subgoals = generate_subgoals(sprint_goal)
        
        # Get stories from the sprint
        stories = get_sprint_stories(jira_client, sprint_id)
        
        # Assign stories to subgoals
        story_assignments = assign_stories_to_subgoals(stories, subgoals)
        
        # Generate achievements for each subgoal
        achievements = generate_achievements(stories, subgoals)
        
        return jsonify({
            'sprint_name': sprint.name,
            'sprint_goal': sprint_goal,
            'subgoals': subgoals,
            'stories': stories,
            'story_assignments': story_assignments,
            'achievements': achievements,
            'start_date': sprint.startDate,
            'end_date': sprint.endDate
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sprint-report/download', methods=['GET'])
def download_sprint_report():
    try:
        board_id = request.args.get('boardId')
        sprint_id = request.args.get('sprintId')
        
        if not board_id or not sprint_id:
            return jsonify({'error': 'Board ID and Sprint ID are required'}), 400

        jira_client = get_jira_client()
        
        # Get sprint details
        sprint = jira_client.sprint(sprint_id)
        if not sprint:
            return jsonify({'error': 'Sprint not found'}), 404
        
        # Get sprint goal
        sprint_goal = sprint.goal if hasattr(sprint, 'goal') else "No sprint goal found"
        
        # Generate subgoals using Gemini
        subgoals = generate_subgoals(sprint_goal)
        
        # Get stories from the sprint
        stories = get_sprint_stories(jira_client, sprint_id)
        
        # Assign stories to subgoals
        story_assignments = assign_stories_to_subgoals(stories, subgoals)
        
        # Generate achievements for each subgoal
        achievements = generate_achievements(stories, subgoals)

        # Create a new Word document
        doc = Document()
        
        # Add title
        title = doc.add_heading(f'Sprint Report: {sprint.name}', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add sprint dates
        dates = doc.add_paragraph()
        dates.alignment = WD_ALIGN_PARAGRAPH.CENTER
        dates.add_run(f'{sprint.startDate} - {sprint.endDate}').italic = True
        
        # Add sprint goal
        doc.add_heading('Sprint Goal', level=1)
        doc.add_paragraph(sprint_goal)
        
        # Add achievements and story assignments
        doc.add_heading('Achievements and Story Assignments', level=1)
        
        # Process achievements and story assignments
        achievement_sections = achievements.split('\n\n')
        assignment_sections = story_assignments.split('\n\n')
        
        for section in achievement_sections:
            if not section.strip():
                continue
                
            lines = section.split('\n')
            if not lines:
                continue
                
            # Add subgoal heading
            subgoal_heading = doc.add_heading(lines[0], level=2)
            
            # Add story numbers if available
            story_numbers = next((line for line in lines if line.startswith('Story Numbers:')), None)
            if story_numbers:
                doc.add_paragraph(story_numbers)
            
            # Add achievements
            for line in lines:
                if line.startswith('- '):
                    p = doc.add_paragraph()
                    p.style = 'List Bullet'
                    p.add_run(line[2:])
        
        # Save the document to a BytesIO object
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        
        return send_file(
            doc_io,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'sprint_report_{sprint.name}.docx'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def process_excel_data(excel_file):
    """Process Excel file and extract relevant information using LLM."""
    try:
        # Read Excel file
        df = pd.read_excel(excel_file)
        
        # Convert DataFrame to string representation
        excel_data = df.to_string()
        
        # Create prompt for LLM to extract structured data
        prompt = f"""
        Analyze the following Excel data and extract the following information in a structured format:
        1. Sprint Capacity
        2. Team Member Capacities
        3. Story Details (including subtasks)
        4. Changelogs
        5. Blockers/Impediments
        
        Excel Data:
        {excel_data}
        
        Return the data in this JSON format:
        {{
            "sprint_capacity": {{
                "total_capacity": number,
                "unit": "hours/days"
            }},
            "team_members": [
                {{
                    "name": string,
                    "capacity": number,
                    "unit": "hours/days"
                }}
            ],
            "stories": [
                {{
                    "id": string,
                    "summary": string,
                    "description": string,
                    "status": string,
                    "assignee": string,
                    "subtasks": [
                        {{
                            "id": string,
                            "summary": string,
                            "status": string,
                            "assignee": string
                        }}
                    ],
                    "changelog": [
                        {{
                            "date": string,
                            "field": string,
                            "from": string,
                            "to": string
                        }}
                    ],
                    "blockers": [
                        {{
                            "description": string,
                            "resolution": string
                        }}
                    ]
                }}
            ]
        }}
        
        Important: Return ONLY the JSON object, with no additional text or explanation.
        """
        
        response = model.generate_content(prompt)
        
        # Clean the response text to ensure it's valid JSON
        response_text = response.text.strip()
        
        # Remove any markdown code block indicators if present
        if response_text.startswith('```json'):
            response_text = response_text[7:]
        if response_text.startswith('```'):
            response_text = response_text[3:]
        if response_text.endswith('```'):
            response_text = response_text[:-3]
            
        response_text = response_text.strip()
        
        try:
            structured_data = json.loads(response_text)
            return structured_data
        except json.JSONDecodeError as e:
            # If JSON parsing fails, try to extract JSON from the response
            import re
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                try:
                    structured_data = json.loads(json_match.group())
                    return structured_data
                except json.JSONDecodeError:
                    raise Exception(f"Failed to parse JSON from response: {str(e)}")
            else:
                raise Exception(f"Failed to extract JSON from response: {str(e)}")
                
    except Exception as e:
        raise Exception(f"Error processing Excel file: {str(e)}")

def parse_jira_datetime(datetime_str):
    """Parse Jira datetime string to datetime object."""
    try:
        if not datetime_str:
            print("Empty datetime string received")
            return None
            
        print(f"Parsing datetime: {datetime_str}")
        
        # Handle different Jira datetime formats
        if '+' in datetime_str:
            # Format: 2025-05-16T15:38:57.738+0530
            dt_str, offset = datetime_str.split('+')
            dt = datetime.fromisoformat(dt_str)
            hours = int(offset[:2])
            minutes = int(offset[2:4])
            dt = dt - timedelta(hours=hours, minutes=minutes)
        elif 'Z' in datetime_str:
            # Format: 2025-05-16T15:38:57.738Z
            dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
        else:
            # Format: 2025-05-16T15:38:57.738
            dt = datetime.fromisoformat(datetime_str)
            
        print(f"Successfully parsed datetime: {dt}")
        return dt
    except Exception as e:
        print(f"Error parsing datetime {datetime_str}: {str(e)}")
        return None

def analyze_sprint_churn(sprint_data):
    """Analyze sprint churn by examining story changes during the sprint."""
    print(f"Analyzing sprint data: {json.dumps(sprint_data, indent=2)}")
    
    # Extract and validate sprint dates
    sprint_start_str = sprint_data.get('start_date')
    sprint_end_str = sprint_data.get('end_date')
    
    if not sprint_start_str or not sprint_end_str:
        print(f"Missing sprint dates. Start: {sprint_start_str}, End: {sprint_end_str}")
        raise Exception("Missing sprint dates")
    
    sprint_start = parse_jira_datetime(sprint_start_str)
    sprint_end = parse_jira_datetime(sprint_end_str)
    
    if not sprint_start:
        print(f"Failed to parse sprint start date: {sprint_start_str}")
        raise Exception(f"Invalid sprint start date: {sprint_start_str}")
    if not sprint_end:
        print(f"Failed to parse sprint end date: {sprint_end_str}")
        raise Exception(f"Invalid sprint end date: {sprint_end_str}")
    
    print(f"Sprint period: {sprint_start} to {sprint_end}")
    
    churn_analysis = {
        'added_stories': [],
        'removed_stories': [],
        'modified_stories': [],
        'status_changes': {},
        'story_point_changes': {},
        'assignee_changes': {}
    }
    
    for story in sprint_data['stories']:
        story_changes = {
            'key': story['key'],
            'summary': story['summary'],
            'changes': [],
            'status_changes': [],
            'point_changes': [],
            'assignee_changes': []
        }
        
        # Analyze changelog entries within sprint dates
        for change in story['changelog']:
            change_date = parse_jira_datetime(change['date'])
            if change_date and sprint_start <= change_date <= sprint_end:
                if change['field'] == 'status':
                    story_changes['status_changes'].append({
                        'date': change['date'],
                        'from': change['from'],
                        'to': change['to']
                    })
                elif change['field'] == 'Story Points':
                    story_changes['point_changes'].append({
                        'date': change['date'],
                        'from': change['from'],
                        'to': change['to']
                    })
                elif change['field'] == 'assignee':
                    story_changes['assignee_changes'].append({
                        'date': change['date'],
                        'from': change['from'],
                        'to': change['to']
                    })
                
                story_changes['changes'].append({
                    'date': change['date'],
                    'field': change['field'],
                    'from': change['from'],
                    'to': change['to']
                })
        
        # Check if story was added during sprint
        story_created = parse_jira_datetime(story['created'])
        if story_created and sprint_start <= story_created <= sprint_end:
            churn_analysis['added_stories'].append({
                'key': story['key'],
                'summary': story['summary'],
                'created_date': story['created']
            })
        
        # If story has changes during sprint, add to modified stories
        if story_changes['changes']:
            churn_analysis['modified_stories'].append(story_changes)
            
            # Track status changes
            if story_changes['status_changes']:
                churn_analysis['status_changes'][story['key']] = story_changes['status_changes']
            
            # Track story point changes
            if story_changes['point_changes']:
                churn_analysis['story_point_changes'][story['key']] = story_changes['point_changes']
            
            # Track assignee changes
            if story_changes['assignee_changes']:
                churn_analysis['assignee_changes'][story['key']] = story_changes['assignee_changes']
    
    return churn_analysis

def generate_improvement_areas(structured_data, sprint_data):
    """Generate improvement areas using LLM based on structured data and sprint data."""
    prompt = f"""
    Analyze the following sprint data and generate detailed improvement areas. Focus on:
    1. Spill-over Analysis
       - Identify stories that spilled over
       - Identify stories from changelog if their sprint number has changed from current sprint to future sprints or removed from current sprint during the sprint.
       - Analyze root causes
       - Suggest preventive measures
    
    2. Churn Analysis
       - Identify stories which were added during the sprint.
       - Identify stories from changelog if their sprint number has changed to current sprint during the sprint.
       - Analyze impact on sprint velocity if there are any churned stories.
       - Suggest ways to reduce churn if there are any churned stories.
    
    3. Team Utilization
       - Calculate utilization for each team member:
         * Get their story point capacity for the sprint
         * Count total story points completed by them
         * Calculate utilization as: (completed story points / capacity) * 100
       - Consider a team member over-utilized if:
         * They complete more story points than their capacity
       - Consider a team member under-utilized if:
         * They handle less story points than their capacity
       - Analyze workload distribution
       - Suggest optimal resource allocation
    
    4. Additional Improvement Areas
       - Identify any other patterns or issues
       - Suggest specific actionable improvements
    
    Sprint Data:
    {json.dumps(sprint_data, indent=2)}
    
    Excel Data:
    {json.dumps(structured_data, indent=2)}
    
    Return the analysis in this JSON format:
    {{
        "spill_over_analysis": {{
            "spilled_stories": [
                {{
                    "story_id": string,
                    "reason": string,
                    "prevention_suggestion": string
                }}
            ],
            "root_causes": [string],
            "recommendations": [string]
        }},
        "churn_analysis": {{
            "high_churn_stories": [
                {{
                    "story_id": string,
                    "churn_count": number,
                    "impact": string
                }}
            ],
            "velocity_impact": string,
            "reduction_suggestions": [string]
        }},
        "team_utilization": {{
            "under_utilized": [
                {{
                    "member": string,
                    "capacity": number,
                    "completed_points": number,
                    "utilization": number,
                    "suggestion": string
                }}
            ],
            "over_utilized": [
                {{
                    "member": string,
                    "capacity": number,
                    "completed_points": number,
                    "utilization": number,
                    "suggestion": string
                }}
            ],
            "workload_distribution": string,
            "optimization_suggestions": [string]
        }},
        "additional_improvements": [
            {{
                "area": string,
                "observation": string,
                "suggestion": string
            }}
        ]
    }}
    
    Important: Return ONLY the JSON object, with no additional text or explanation.
    """
    
    response = model.generate_content(prompt)
    
    # Clean the response text to ensure it's valid JSON
    response_text = response.text.strip()
    
    # Remove any markdown code block indicators if present
    if response_text.startswith('```json'):
        response_text = response_text[7:]
    if response_text.startswith('```'):
        response_text = response_text[3:]
    if response_text.endswith('```'):
        response_text = response_text[:-3]
        
    response_text = response_text.strip()
    
    try:
        return json.loads(response_text)
    except json.JSONDecodeError as e:
        # If JSON parsing fails, try to extract JSON from the response
        import re
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group())
            except json.JSONDecodeError:
                raise Exception(f"Failed to parse JSON from response: {str(e)}")
        else:
            raise Exception(f"Failed to extract JSON from response: {str(e)}")

def generate_sprint_analysis_doc(sprint_data, improvement_areas):
    """Generate a Word document containing sprint analysis."""
    doc = Document()
    
    # Add title
    title = doc.add_heading(f'Sprint Analysis Report: {sprint_data["sprint_name"]}', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add sprint details
    doc.add_heading('Sprint Details', level=1)
    doc.add_paragraph(f'Sprint Goal: {sprint_data["sprint_goal"] or "No sprint goal defined"}')
    doc.add_paragraph(f'Start Date: {sprint_data["start_date"]}')
    doc.add_paragraph(f'End Date: {sprint_data["end_date"]}')
    
    # Add Spill-over Analysis
    doc.add_heading('Spill-over Analysis', level=1)
    if improvement_areas['spill_over_analysis']['spilled_stories']:
        doc.add_paragraph('Spilled Stories:')
        for story in improvement_areas['spill_over_analysis']['spilled_stories']:
            p = doc.add_paragraph()
            p.add_run(f'Story ID: {story["story_id"]}\n').bold = True
            p.add_run(f'Reason: {story["reason"]}\n')
            p.add_run(f'Prevention Suggestion: {story["prevention_suggestion"]}')
    else:
        doc.add_paragraph('No stories spilled over in this sprint.')
    
    doc.add_paragraph('Root Causes:')
    for cause in improvement_areas['spill_over_analysis']['root_causes']:
        doc.add_paragraph(cause, style='List Bullet')
    
    doc.add_paragraph('Recommendations:')
    for rec in improvement_areas['spill_over_analysis']['recommendations']:
        doc.add_paragraph(rec, style='List Bullet')
    
    # Add Churn Analysis
    doc.add_heading('Churn Analysis', level=1)
    if improvement_areas['churn_analysis']['high_churn_stories']:
        doc.add_paragraph('High Churn Stories:')
        for story in improvement_areas['churn_analysis']['high_churn_stories']:
            p = doc.add_paragraph()
            p.add_run(f'Story ID: {story["story_id"]}\n').bold = True
            p.add_run(f'Churn Count: {story["churn_count"]}\n')
            p.add_run(f'Impact: {story["impact"]}')
    else:
        doc.add_paragraph('No high churn stories identified in this sprint.')
    
    doc.add_paragraph(f'Velocity Impact: {improvement_areas["churn_analysis"]["velocity_impact"]}')
    
    doc.add_paragraph('Reduction Suggestions:')
    for suggestion in improvement_areas['churn_analysis']['reduction_suggestions']:
        doc.add_paragraph(suggestion, style='List Bullet')
    
    # Add Team Utilization
    doc.add_heading('Team Utilization', level=1)
    
    doc.add_heading('Over-Utilized Team Members', level=2)
    if improvement_areas['team_utilization']['over_utilized']:
        for member in improvement_areas['team_utilization']['over_utilized']:
            p = doc.add_paragraph()
            p.add_run(f'Member: {member["member"]}\n').bold = True
            p.add_run(f'Capacity: {member["capacity"]} points\n')
            p.add_run(f'Completed Points: {member["completed_points"]}\n')
            p.add_run(f'Utilization: {member["utilization"]}%\n')
            p.add_run(f'Suggestion: {member["suggestion"]}')
    else:
        doc.add_paragraph('No over-utilized team members identified.')
    
    doc.add_heading('Under-Utilized Team Members', level=2)
    if improvement_areas['team_utilization']['under_utilized']:
        for member in improvement_areas['team_utilization']['under_utilized']:
            p = doc.add_paragraph()
            p.add_run(f'Member: {member["member"]}\n').bold = True
            p.add_run(f'Capacity: {member["capacity"]} points\n')
            p.add_run(f'Completed Points: {member["completed_points"]}\n')
            p.add_run(f'Utilization: {member["utilization"]}%\n')
            p.add_run(f'Suggestion: {member["suggestion"]}')
    else:
        doc.add_paragraph('No under-utilized team members identified.')
    
    doc.add_paragraph(f'Workload Distribution: {improvement_areas["team_utilization"]["workload_distribution"]}')
    
    doc.add_paragraph('Optimization Suggestions:')
    for suggestion in improvement_areas['team_utilization']['optimization_suggestions']:
        doc.add_paragraph(suggestion, style='List Bullet')
    
    # Add Additional Improvements
    doc.add_heading('Additional Improvements', level=1)
    for improvement in improvement_areas['additional_improvements']:
        p = doc.add_paragraph()
        p.add_run(f'Area: {improvement["area"]}\n').bold = True
        p.add_run(f'Observation: {improvement["observation"]}\n')
        p.add_run(f'Suggestion: {improvement["suggestion"]}')
    
    return doc

@app.route('/api/sprint-analysis/download', methods=['POST'])
def download_sprint_analysis():
    try:
        data = request.get_json()
        if not data or 'sprint_data' not in data or 'improvement_areas' not in data:
            return jsonify({'error': 'Invalid request data'}), 400
        
        # Generate Word document
        doc = generate_sprint_analysis_doc(data['sprint_data'], data['improvement_areas'])
        
        # Save to BytesIO
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        
        return send_file(
            doc_io,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'sprint_analysis_{data["sprint_data"]["sprint_name"]}.docx'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sprint-analysis', methods=['POST'])
def analyze_sprint():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        board_id = request.form.get('boardId')
        sprint_id = request.form.get('sprintId')
        
        if not board_id or not sprint_id:
            return jsonify({'error': 'Board ID and Sprint ID are required'}), 400
        
        excel_file = request.files['file']
        if not excel_file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({'error': 'Invalid file format. Please upload an Excel file.'}), 400
        
        # Get sprint details
        jira_client = get_jira_client()
        sprint = jira_client.sprint(sprint_id)
        if not sprint:
            return jsonify({'error': 'Sprint not found'}), 404
        
        # Get sprint stories with all details
        sprint_stories = get_sprint_stories(jira_client, sprint_id)
        
        # Process Excel data
        structured_data = process_excel_data(excel_file)
        
        # Generate improvement areas using both Excel data and sprint stories
        sprint_data = {
            'sprint_name': sprint.name,
            'sprint_goal': sprint.goal if hasattr(sprint, 'goal') else None,
            'start_date': sprint.startDate,
            'end_date': sprint.endDate,
            'stories': sprint_stories
        }
        
        improvement_areas = generate_improvement_areas(structured_data, sprint_data)
        
        return jsonify({
            'sprint_name': sprint.name,
            'sprint_goal': sprint.goal if hasattr(sprint, 'goal') else None,
            'start_date': sprint.startDate,
            'end_date': sprint.endDate,
            'structured_data': structured_data,
            'improvement_areas': improvement_areas
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True) 