import pandas as pd
from datetime import datetime, date
from typing import Dict, List, Optional
from data_manager import DataManager

class EVMCalculator:
    def __init__(self, data_manager: DataManager):
        self.data_manager = data_manager
    
    def calculate_project_kpi(self, project_name: str) -> Optional[Dict]:
        """Calculate EVM KPIs for a specific project"""
        try:
            # Get project info
            project_info = self.data_manager.get_project_info(project_name)
            if not project_info:
                return None
            
            # Get latest progress data
            progress_data = self.data_manager.get_progress_data(project_name)
            if progress_data.empty:
                return None
            
            # Get the latest entry
            latest_data = progress_data.iloc[-1]
            
            # Calculate PV, EV, AC
            total_budget = project_info['total_budget']
            planned_completion = latest_data['planned_completion'] / 100
            actual_completion = latest_data['actual_completion'] / 100
            
            pv = total_budget * planned_completion  # Planned Value
            ev = total_budget * actual_completion   # Earned Value
            ac = latest_data['actual_cost']         # Actual Cost
            
            # Calculate KPIs
            cpi = ev / ac if ac > 0 else 0          # Cost Performance Index
            spi = ev / pv if pv > 0 else 0          # Schedule Performance Index
            cv = ev - ac                            # Cost Variance
            sv = ev - pv                            # Schedule Variance
            
            # Calculate progress percentages
            cost_variance_percent = (cv / pv * 100) if pv > 0 else 0
            schedule_variance_percent = (sv / pv * 100) if pv > 0 else 0
            
            # Determine project status
            status = self._determine_project_status(spi, cpi)
            
            # Calculate completion estimates
            eac = self._calculate_eac(total_budget, cpi, actual_completion)  # Estimate at Completion
            etc = eac - ac if eac > ac else 0                              # Estimate to Complete
            
            return {
                'project_name': project_name,
                'pv': pv,
                'ev': ev,
                'ac': ac,
                'cpi': cpi,
                'spi': spi,
                'cv': cv,
                'sv': sv,
                'cost_variance_percent': cost_variance_percent,
                'schedule_variance_percent': schedule_variance_percent,
                'status': status,
                'eac': eac,
                'etc': etc,
                'planned_completion': planned_completion * 100,
                'actual_completion': actual_completion * 100,
                'total_budget': total_budget
            }
        except Exception as e:
            print(f"Error calculating project KPI: {e}")
            return None
    
    def calculate_portfolio_kpi(self) -> Optional[Dict]:
        """Calculate aggregated KPIs for the entire portfolio"""
        try:
            projects = self.data_manager.get_all_projects()
            if not projects:
                return None
            
            total_pv = 0
            total_ev = 0
            total_ac = 0
            valid_projects = 0
            project_details = []
            
            for project in projects:
                project_kpi = self.calculate_project_kpi(project['project_name'])
                if project_kpi:
                    total_pv += project_kpi['pv']
                    total_ev += project_kpi['ev']
                    total_ac += project_kpi['ac']
                    valid_projects += 1
                    project_details.append(project_kpi)
            
            if valid_projects == 0:
                return None
            
            # Calculate portfolio-level KPIs
            avg_cpi = total_ev / total_ac if total_ac > 0 else 0
            avg_spi = total_ev / total_pv if total_pv > 0 else 0
            total_cv = total_ev - total_ac
            total_sv = total_ev - total_pv
            
            # Count projects by status
            status_counts = {'متقدم': 0, 'متأخر': 0, 'على المسار': 0}
            for project in project_details:
                status = project['status']
                if status in status_counts:
                    status_counts[status] += 1
            
            return {
                'total_projects': valid_projects,
                'total_pv': total_pv,
                'total_ev': total_ev,
                'total_ac': total_ac,
                'avg_cpi': avg_cpi,
                'avg_spi': avg_spi,
                'total_cv': total_cv,
                'total_sv': total_sv,
                'status_counts': status_counts,
                'project_details': project_details
            }
        except Exception as e:
            print(f"Error calculating portfolio KPI: {e}")
            return None
    
    def get_all_projects_performance(self) -> List[Dict]:
        """Get performance data for all projects"""
        try:
            projects = self.data_manager.get_all_projects()
            performance_data = []
            
            for project in projects:
                project_kpi = self.calculate_project_kpi(project['project_name'])
                if project_kpi:
                    performance_data.append({
                        'project_name': project['project_name'],
                        'cpi': project_kpi['cpi'],
                        'spi': project_kpi['spi'],
                        'status': project_kpi['status'],
                        'completion': project_kpi['actual_completion']
                    })
            
            return performance_data
        except Exception as e:
            print(f"Error getting projects performance: {e}")
            return []
    
    def get_dashboard_data(self, status_filter: str, spi_threshold: float, cpi_threshold: float) -> List[Dict]:
        """Get filtered data for dashboard"""
        try:
            projects = self.data_manager.get_all_projects()
            dashboard_data = []
            
            for project in projects:
                project_kpi = self.calculate_project_kpi(project['project_name'])
                if project_kpi:
                    # Apply filters
                    include_project = True
                    
                    if status_filter != "جميع المشاريع":
                        if project_kpi['status'] != status_filter:
                            include_project = False
                    
                    # Apply SPI and CPI thresholds
                    if project_kpi['spi'] < spi_threshold or project_kpi['cpi'] < cpi_threshold:
                        if status_filter == "على المسار":
                            include_project = False
                    
                    if include_project:
                        dashboard_data.append(project_kpi)
            
            return dashboard_data
        except Exception as e:
            print(f"Error getting dashboard data: {e}")
            return []
    
    def _determine_project_status(self, spi: float, cpi: float) -> str:
        """Determine project status based on SPI and CPI"""
        if spi >= 1.0 and cpi >= 1.0:
            return "متقدم"
        elif spi >= 0.9 and cpi >= 0.9:
            return "على المسار"
        else:
            return "متأخر"
    
    def _calculate_eac(self, total_budget: float, cpi: float, completion_percent: float) -> float:
        """Calculate Estimate at Completion"""
        if cpi > 0:
            return total_budget / cpi
        else:
            return total_budget
    
    def calculate_trend_analysis(self, project_name: str) -> Optional[Dict]:
        """Calculate trend analysis for a project"""
        try:
            progress_data = self.data_manager.get_progress_data(project_name)
            if progress_data.empty or len(progress_data) < 2:
                return None
            
            # Sort by date
            progress_data['entry_date'] = pd.to_datetime(progress_data['entry_date'])
            progress_data = progress_data.sort_values('entry_date')
            
            # Calculate trends
            latest_cpi = None
            latest_spi = None
            cpi_trend = []
            spi_trend = []
            
            project_info = self.data_manager.get_project_info(project_name)
            total_budget = project_info['total_budget'] if project_info else 0
            
            for _, row in progress_data.iterrows():
                if total_budget > 0:
                    planned_completion = row['planned_completion'] / 100
                    actual_completion = row['actual_completion'] / 100
                    
                    pv = total_budget * planned_completion
                    ev = total_budget * actual_completion
                    ac = row['actual_cost']
                    
                    cpi = ev / ac if ac > 0 else 0
                    spi = ev / pv if pv > 0 else 0
                    
                    cpi_trend.append(cpi)
                    spi_trend.append(spi)
                    
                    latest_cpi = cpi
                    latest_spi = spi
            
            # Calculate trend direction
            cpi_direction = self._calculate_trend_direction(cpi_trend)
            spi_direction = self._calculate_trend_direction(spi_trend)
            
            return {
                'latest_cpi': latest_cpi,
                'latest_spi': latest_spi,
                'cpi_trend': cpi_direction,
                'spi_trend': spi_direction,
                'data_points': len(progress_data)
            }
        except Exception as e:
            print(f"Error calculating trend analysis: {e}")
            return None
    
    def _calculate_trend_direction(self, values: List[float]) -> str:
        """Calculate trend direction from a list of values"""
        if len(values) < 2:
            return "مستقر"
        
        recent_values = values[-3:] if len(values) >= 3 else values
        
        if len(recent_values) < 2:
            return "مستقر"
        
        # Calculate average change
        changes = []
        for i in range(1, len(recent_values)):
            changes.append(recent_values[i] - recent_values[i-1])
        
        avg_change = sum(changes) / len(changes)
        
        if avg_change > 0.05:
            return "تحسن"
        elif avg_change < -0.05:
            return "تراجع"
        else:
            return "مستقر"
