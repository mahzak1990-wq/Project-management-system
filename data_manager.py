import pandas as pd
import os
import sqlite3
from datetime import datetime
import zipfile
import shutil
from typing import Dict, List, Optional
import json

class DataManager:
    def __init__(self):
        self.data_dir = "data"
        self.db_path = os.path.join(self.data_dir, "projects.db")
        self.backup_dir = "backups"
        self.ensure_directories()
        self.init_database()
        self.migrate_database()
    
    def ensure_directories(self):
        """Create necessary directories if they don't exist"""
        for directory in [self.data_dir, self.backup_dir]:
            os.makedirs(directory, exist_ok=True)
    
    def init_database(self):
        """Initialize SQLite database with required tables"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Parent categories table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS parent_categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    category_name TEXT UNIQUE NOT NULL,
                    description TEXT,
                    created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Projects table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS projects (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_name TEXT UNIQUE NOT NULL,
                    project_id TEXT,
                    purchase_order TEXT,
                    parent_category_id INTEGER DEFAULT NULL,
                    executing_company TEXT,
                    consulting_company TEXT,
                    start_date DATE,
                    end_date DATE,
                    total_budget REAL,
                    project_location TEXT,
                    project_type TEXT,
                    project_description TEXT,
                    display_order INTEGER DEFAULT 0,
                    created_date TIMESTAMP,
                    FOREIGN KEY (parent_category_id) REFERENCES parent_categories (id)
                )
            ''')
            
            # Progress data table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS progress_data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_name TEXT,
                    entry_date DATE,
                    planned_completion REAL,
                    planned_cost REAL,
                    actual_completion REAL,
                    actual_cost REAL,
                    notes TEXT,
                    FOREIGN KEY (project_name) REFERENCES projects (project_name)
                )
            ''')
            
            # Resources table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS resources (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_name TEXT,
                    resource_type TEXT,
                    name TEXT,
                    quantity INTEGER,
                    daily_rate REAL,
                    start_date DATE,
                    end_date DATE,
                    notes TEXT,
                    FOREIGN KEY (project_name) REFERENCES projects (project_name)
                )
            ''')
            
            # Original Excel files table to store imported files exactly as they are
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS original_excel_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_name TEXT NOT NULL,
                    file_content BLOB NOT NULL,
                    imported_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    projects_imported TEXT,
                    file_hash TEXT UNIQUE
                )
            ''')
            
            # Insert default parent categories if they don't exist
            default_categories = [
                ('Uncategorized', 'Projects not yet assigned to a category'),
                ('Sewerage Projects', 'Water and sewerage infrastructure projects'),
                ('Water Projects', 'Water supply and distribution projects'),  
                ('Construction Projects', 'General construction and building projects')
            ]
            
            for cat_name, cat_desc in default_categories:
                cursor.execute('''
                    INSERT OR IGNORE INTO parent_categories (category_name, description)
                    VALUES (?, ?)
                ''', (cat_name, cat_desc))
            
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Database initialization error: {e}")
    
    def migrate_database(self):
        """Migrate database schema to add new columns if they don't exist"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if new columns exist and add them if they don't
            cursor.execute("PRAGMA table_info(projects)")
            columns = [column[1] for column in cursor.fetchall()]
            
            if 'project_id' not in columns:
                cursor.execute('ALTER TABLE projects ADD COLUMN project_id TEXT')
                print("Added project_id column to projects table")
            
            if 'parent_category_id' not in columns:
                cursor.execute('ALTER TABLE projects ADD COLUMN parent_category_id INTEGER DEFAULT NULL')
                print("Added parent_category_id column to projects table")
                
            if 'display_order' not in columns:
                cursor.execute('ALTER TABLE projects ADD COLUMN display_order INTEGER DEFAULT 0')
                print("Added display_order column to projects table")
                
            if 'contractor_name' not in columns:
                cursor.execute('ALTER TABLE projects ADD COLUMN contractor_name TEXT')
                print("Added contractor_name column to projects table")
                
            if 'project_manager' not in columns:
                cursor.execute('ALTER TABLE projects ADD COLUMN project_manager TEXT')
                print("Added project_manager column to projects table")
            
            conn.commit()
            conn.close()
            print("Database migration completed successfully")
            
        except Exception as e:
            print(f"Database migration error: {e}")
    
    def add_project(self, project_data: Dict) -> bool:
        """Add a new project to the database"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT OR REPLACE INTO projects 
                (project_name, project_id, parent_category_id, executing_company, consulting_company, start_date, 
                 end_date, total_budget, project_location, project_type, 
                 project_description, display_order, contractor_name, project_manager, created_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                project_data['project_name'],
                project_data.get('project_id', ''),
                project_data.get('parent_category_id', None),
                project_data['executing_company'],
                project_data['consulting_company'],
                project_data['start_date'],
                project_data['end_date'],
                project_data['total_budget'],
                project_data['project_location'],
                project_data['project_type'],
                project_data['project_description'],
                project_data.get('display_order', 0),
                project_data.get('contractor_name', ''),
                project_data.get('project_manager', ''),
                project_data['created_date']
            ))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding project: {e}")
            return False
    
    def get_all_projects(self) -> List[Dict]:
        """Retrieve all projects from the database with parent category info"""
        try:
            conn = sqlite3.connect(self.db_path)
            query = '''
                SELECT p.*, pc.category_name as parent_category_name, pc.description as parent_category_description
                FROM projects p
                LEFT JOIN parent_categories pc ON p.parent_category_id = pc.id
                ORDER BY pc.category_name, p.display_order, p.created_date DESC
            '''
            df = pd.read_sql_query(query, conn)
            conn.close()
            return df.to_dict('records')
        except Exception as e:
            print(f"Error retrieving projects: {e}")
            return []
    
    def get_parent_categories(self) -> List[Dict]:
        """Get all parent categories"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM parent_categories ORDER BY category_name")
            categories = []
            for row in cursor.fetchall():
                categories.append({
                    'id': row[0],
                    'category_name': row[1], 
                    'description': row[2],
                    'created_date': row[3]
                })
            conn.close()
            return categories
        except Exception as e:
            print(f"Error getting parent categories: {e}")
            return []
    
    def update_project_parent_category(self, project_name: str, new_parent_category_id: int, new_display_order: int = 0) -> bool:
        """Move a project to a different parent category"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE projects 
                SET parent_category_id = ?, display_order = ?
                WHERE project_name = ?
            ''', (new_parent_category_id, new_display_order, project_name))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error updating project parent category: {e}")
            return False
    
    def get_projects_by_category(self) -> Dict:
        """Get projects grouped by parent category"""
        try:
            all_projects = self.get_all_projects()
            grouped = {}
            
            for project in all_projects:
                category_name = project.get('parent_category_name', 'Uncategorized')
                if category_name not in grouped:
                    grouped[category_name] = []
                grouped[category_name].append(project)
            
            return grouped
        except Exception as e:
            print(f"Error grouping projects by category: {e}")
            return {}
    
    def get_project_info(self, project_name: str) -> Optional[Dict]:
        """Get detailed information for a specific project"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM projects WHERE project_name = ?", (project_name,))
            result = cursor.fetchone()
            conn.close()
            
            if result:
                columns = ['id', 'project_name', 'executing_company', 'consulting_company',
                          'start_date', 'end_date', 'total_budget', 'project_location',
                          'project_type', 'project_description', 'created_date']
                return dict(zip(columns, result))
            return None
        except Exception as e:
            print(f"Error retrieving project info: {e}")
            return None
    
    def get_project_by_name(self, project_name: str) -> Optional[Dict]:
        """Get project by name - alias for get_project_info"""
        return self.get_project_info(project_name)
    
    def add_progress_data(self, progress_data: Dict) -> bool:
        """Add progress data entry"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO progress_data 
                (project_name, entry_date, planned_completion, planned_cost,
                 actual_completion, actual_cost, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                progress_data['project_name'],
                progress_data['entry_date'],
                progress_data['planned_completion'],
                progress_data['planned_cost'],
                progress_data['actual_completion'],
                progress_data['actual_cost'],
                progress_data['notes']
            ))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding progress data: {e}")
            return False
    
    def get_progress_data(self, project_name: str) -> pd.DataFrame:
        """Retrieve progress data for a specific project"""
        try:
            conn = sqlite3.connect(self.db_path)
            df = pd.read_sql_query(
                "SELECT entry_date, planned_completion, planned_cost, actual_completion, actual_cost, notes FROM progress_data WHERE project_name = ? ORDER BY entry_date",
                conn,
                params=[project_name]
            )
            conn.close()
            return df
        except Exception as e:
            print(f"Error retrieving progress data: {e}")
            return pd.DataFrame()
    
    def delete_project_progress(self, project_name: str) -> bool:
        """Delete only progress data for a project (for updates)"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Delete progress data only
            cursor.execute('DELETE FROM progress_data WHERE project_name = ?', (project_name,))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error deleting project progress data: {e}")
            return False

    def delete_project(self, project_name: str) -> bool:
        """Delete a project and all its related data"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Delete progress data first (foreign key constraint)
            cursor.execute('DELETE FROM progress_data WHERE project_name = ?', (project_name,))
            
            # Delete cash flow data
            cursor.execute('DELETE FROM cash_flow WHERE project_name = ?', (project_name,))
            
            # Delete resources
            cursor.execute('DELETE FROM resources WHERE project_name = ?', (project_name,))
            
            # Delete the project itself
            cursor.execute('DELETE FROM projects WHERE project_name = ?', (project_name,))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error deleting project: {e}")
            return False
    
    def add_resource(self, resource_data: Dict) -> bool:
        """Add resource (labor or equipment)"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO resources 
                (project_name, resource_type, name, quantity, daily_rate,
                 start_date, end_date, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                resource_data['project_name'],
                resource_data['resource_type'],
                resource_data['name'],
                resource_data['quantity'],
                resource_data['daily_rate'],
                resource_data['start_date'],
                resource_data['end_date'],
                resource_data['notes']
            ))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding resource: {e}")
            return False
    
    def get_resources(self, project_name: str, resource_type: str = None) -> pd.DataFrame:
        """Retrieve resources for a specific project"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            if resource_type:
                df = pd.read_sql_query(
                    "SELECT * FROM resources WHERE project_name = ? AND resource_type = ?",
                    conn,
                    params=[project_name, resource_type]
                )
            else:
                df = pd.read_sql_query(
                    "SELECT * FROM resources WHERE project_name = ?",
                    conn,
                    params=[project_name]
                )
            
            conn.close()
            return df
        except Exception as e:
            print(f"Error retrieving resources: {e}")
            return pd.DataFrame()
    
    def get_cash_flow_data(self, project_name: str = None, start_date=None, end_date=None) -> pd.DataFrame:
        """Get cash flow data for reporting with proper date filtering"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            if project_name:
                query = """
                    SELECT p.entry_date, p.planned_cost, p.actual_cost, p.planned_completion, 
                           p.actual_completion, pr.project_name, pr.total_budget
                    FROM progress_data p
                    JOIN projects pr ON p.project_name = pr.project_name
                    WHERE p.project_name = ?
                """
                params = [project_name]
            else:
                query = """
                    SELECT p.entry_date, p.planned_cost, p.actual_cost, p.planned_completion,
                           p.actual_completion, pr.project_name, pr.total_budget
                    FROM progress_data p
                    JOIN projects pr ON p.project_name = pr.project_name
                """
                params = []
            
            if start_date and end_date:
                query += " AND p.entry_date BETWEEN ? AND ?"
                params.extend([start_date, end_date])
            
            query += " ORDER BY pr.project_name, p.entry_date"
            
            df = pd.read_sql_query(query, conn, params=params)
            conn.close()
            
            # Convert entry_date to datetime for better processing
            if not df.empty:
                df['entry_date'] = pd.to_datetime(df['entry_date'])
            
            return df
        except Exception as e:
            print(f"Error retrieving cash flow data: {e}")
            return pd.DataFrame()
    
    def get_aggregated_financial_data(self, project_name: str = None, start_date=None, end_date=None, 
                                    aggregation_type='daily') -> pd.DataFrame:
        """Get aggregated financial data by day, month, or year from imported Excel data"""
        try:
            cash_flow_data = self.get_cash_flow_data(project_name, start_date, end_date)
            
            if cash_flow_data.empty:
                return pd.DataFrame()
            
            # Group by the specified aggregation type
            if aggregation_type == 'daily':
                cash_flow_data['period'] = cash_flow_data['entry_date'].dt.date
            elif aggregation_type == 'monthly':
                cash_flow_data['period'] = cash_flow_data['entry_date'].dt.to_period('M')
            elif aggregation_type == 'yearly':
                cash_flow_data['period'] = cash_flow_data['entry_date'].dt.to_period('Y')
            
            # Aggregate by period and project
            aggregated = cash_flow_data.groupby(['project_name', 'period']).agg({
                'planned_cost': 'sum',
                'actual_cost': 'sum',
                'planned_completion': 'mean',
                'actual_completion': 'mean',
                'total_budget': 'first'
            }).reset_index()
            
            return aggregated
            
        except Exception as e:
            print(f"Error getting aggregated financial data: {e}")
            return pd.DataFrame()
    
    def create_backup(self) -> bool:
        """Create a backup of all data"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"backup_{timestamp}.zip"
            backup_path = os.path.join(self.backup_dir, backup_filename)
            
            with zipfile.ZipFile(backup_path, 'w') as backup_zip:
                # Add database file
                if os.path.exists(self.db_path):
                    backup_zip.write(self.db_path, "projects.db")
                
                # Add any CSV files if they exist
                for file in os.listdir(self.data_dir):
                    if file.endswith('.csv'):
                        file_path = os.path.join(self.data_dir, file)
                        backup_zip.write(file_path, file)
            
            return True
        except Exception as e:
            print(f"Backup error: {e}")
            return False
    
    def restore_backup(self, uploaded_file) -> bool:
        """Restore data from backup file"""
        try:
            # Save uploaded file temporarily
            temp_path = os.path.join(self.backup_dir, "temp_restore.zip")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.read())
            
            # Extract backup
            with zipfile.ZipFile(temp_path, 'r') as backup_zip:
                backup_zip.extractall(self.data_dir)
            
            # Clean up temporary file
            os.remove(temp_path)
            
            # Reinitialize database connection
            self.init_database()
            
            return True
        except Exception as e:
            print(f"Restore error: {e}")
            return False
    
    def clear_all_data(self) -> bool:
        """Clear all data from the database"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("DELETE FROM progress_data")
            cursor.execute("DELETE FROM resources")
            cursor.execute("DELETE FROM projects")
            cursor.execute("DELETE FROM original_excel_files")
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error clearing data: {e}")
            return False
    
    def get_data_statistics(self) -> Dict:
        """Get statistics about stored data"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Count projects
            cursor.execute("SELECT COUNT(*) FROM projects")
            total_projects = cursor.fetchone()[0]
            
            # Count total records
            cursor.execute("SELECT COUNT(*) FROM progress_data")
            progress_records = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM resources")
            resource_records = cursor.fetchone()[0]
            
            total_records = progress_records + resource_records
            
            # Calculate data size
            data_size = 0
            if os.path.exists(self.db_path):
                data_size = os.path.getsize(self.db_path) / (1024 * 1024)  # Convert to MB
            
            conn.close()
            
            return {
                'total_projects': total_projects,
                'total_records': total_records,
                'data_size': data_size
            }
        except Exception as e:
            print(f"Error getting statistics: {e}")
            return {}
    
    def save_original_excel_file(self, file_name: str, file_content: bytes, projects_imported: List, file_hash: str) -> bool:
        """Save original Excel file exactly as imported"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Convert projects list to string - handle both string and dict formats
            if projects_imported:
                # If list contains dictionaries, extract project names
                if isinstance(projects_imported[0], dict):
                    project_names = [proj.get('project_name', '') for proj in projects_imported]
                    projects_str = ','.join(project_names)
                else:
                    # If list contains strings, join directly
                    projects_str = ','.join(projects_imported)
            else:
                projects_str = ''
            
            cursor.execute('''
                INSERT OR REPLACE INTO original_excel_files 
                (file_name, file_content, projects_imported, file_hash)
                VALUES (?, ?, ?, ?)
            ''', (file_name, file_content, projects_str, file_hash))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error saving original Excel file: {e}")
            return False
    
    def get_latest_original_excel_file(self) -> Dict:
        """Get the most recently imported original Excel file"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                SELECT file_name, file_content, imported_date, projects_imported 
                FROM original_excel_files 
                ORDER BY imported_date DESC 
                LIMIT 1
            ''')
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {
                    'file_name': result[0],
                    'file_content': result[1],
                    'imported_date': result[2],
                    'projects_imported': result[3].split(',') if result[3] else []
                }
            return {}
        except Exception as e:
            print(f"Error getting original Excel file: {e}")
            return {}
    
    def clear_original_excel_files(self) -> bool:
        """Clear all saved original Excel files"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM original_excel_files")
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error clearing original Excel files: {e}")
            return False
