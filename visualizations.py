import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import streamlit as st
from datetime import datetime, date
from typing import List, Dict, Optional
from data_manager import DataManager
from evm_calculator import EVMCalculator

def create_s_curve(data_manager: DataManager, project_names: List[str]) -> Optional[go.Figure]:
    """Create S-curve visualization for cost progression"""
    try:
        fig = go.Figure()
        
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f']
        
        for idx, project_name in enumerate(project_names):
            # Get progress data
            progress_data = data_manager.get_progress_data(project_name)
            if progress_data.empty:
                continue
            
            # Get project info for budget
            project_info = data_manager.get_project_info(project_name)
            if not project_info:
                continue
            
            total_budget = project_info['total_budget']
            
            # Sort by date
            progress_data['entry_date'] = pd.to_datetime(progress_data['entry_date'])
            progress_data = progress_data.sort_values('entry_date')
            
            # Calculate cumulative values
            planned_values = []
            actual_values = []
            dates = []
            
            for _, row in progress_data.iterrows():
                planned_cost = total_budget * (row['planned_completion'] / 100)
                actual_cost = row['actual_cost']
                
                planned_values.append(planned_cost)
                actual_values.append(actual_cost)
                dates.append(row['entry_date'])
            
            color = colors[idx % len(colors)]
            
            # Add planned curve
            fig.add_trace(go.Scatter(
                x=dates,
                y=planned_values,
                mode='lines+markers',
                name=f'{project_name} - مخطط',
                line=dict(color=color, dash='solid'),
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             'التاريخ: %{x}<br>' +
                             'القيمة: %{y:,.0f} ريال<extra></extra>'
            ))
            
            # Add actual curve
            fig.add_trace(go.Scatter(
                x=dates,
                y=actual_values,
                mode='lines+markers',
                name=f'{project_name} - فعلي',
                line=dict(color=color, dash='dash'),
                hovertemplate='<b>%{fullData.name}</b><br>' +
                             'التاريخ: %{x}<br>' +
                             'القيمة: %{y:,.0f} ريال<extra></extra>'
            ))
        
        # Update layout
        fig.update_layout(
            title={
                'text': 'منحنى S للتكلفة - مقارنة المخطط مقابل الفعلي',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 16}
            },
            xaxis_title='التاريخ',
            yaxis_title='التكلفة التراكمية (ريال)',
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            font=dict(family="Arial", size=12),
            height=500
        )
        
        # Format y-axis
        fig.update_traces(yaxis='y')
        fig.update_layout(yaxis=dict(tickformat=',.0f'))
        
        return fig
    except Exception as e:
        print(f"Error creating S-curve: {e}")
        return None

def create_kpi_dashboard(dashboard_data: List[Dict]):
    """Create KPI dashboard visualization"""
    try:
        if not dashboard_data:
            st.warning("لا توجد بيانات للعرض")
            return
        
        # Convert to DataFrame for easier processing
        df = pd.DataFrame(dashboard_data)
        
        # Key metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            avg_cpi = df['cpi'].mean()
            st.metric(
                "متوسط مؤشر التكلفة (CPI)",
                f"{avg_cpi:.3f}",
                delta=f"{avg_cpi - 1:.3f}" if avg_cpi != 1 else None
            )
        
        with col2:
            avg_spi = df['spi'].mean()
            st.metric(
                "متوسط مؤشر الجدولة (SPI)",
                f"{avg_spi:.3f}",
                delta=f"{avg_spi - 1:.3f}" if avg_spi != 1 else None
            )
        
        with col3:
            total_budget = df['total_budget'].sum()
            st.metric(
                "إجمالي الميزانية",
                f"{total_budget:,.0f} ريال"
            )
        
        with col4:
            avg_completion = df['actual_completion'].mean()
            st.metric(
                "متوسط نسبة الإنجاز",
                f"{avg_completion:.1f}%"
            )
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            # CPI vs SPI scatter plot
            fig_scatter = px.scatter(
                df,
                x='cpi',
                y='spi',
                color='status',
                size='total_budget',
                hover_name='project_name',
                title='مؤشر التكلفة مقابل مؤشر الجدولة',
                labels={
                    'cpi': 'مؤشر أداء التكلفة (CPI)',
                    'spi': 'مؤشر أداء الجدولة (SPI)',
                    'status': 'الحالة',
                    'total_budget': 'الميزانية'
                }
            )
            
            # Add reference lines
            fig_scatter.add_hline(y=1, line_dash="dash", line_color="gray", annotation_text="SPI = 1.0")
            fig_scatter.add_vline(x=1, line_dash="dash", line_color="gray", annotation_text="CPI = 1.0")
            
            fig_scatter.update_layout(height=400)
            st.plotly_chart(fig_scatter, use_container_width=True)
        
        with col2:
            # Project status distribution
            status_counts = df['status'].value_counts()
            
            fig_pie = px.pie(
                values=status_counts.values,
                names=status_counts.index,
                title='توزيع حالات المشاريع'
            )
            
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)
        
        # Detailed table
        st.subheader("تفاصيل المشاريع")
        
        # Format DataFrame for display
        display_df = df.copy()
        display_df = display_df[['project_name', 'cpi', 'spi', 'actual_completion', 'status', 'total_budget']]
        display_df.columns = ['اسم المشروع', 'مؤشر CPI', 'مؤشر SPI', 'نسبة الإنجاز (%)', 'الحالة', 'الميزانية الإجمالية']
        display_df['مؤشر CPI'] = display_df['مؤشر CPI'].round(3)
        display_df['مؤشر SPI'] = display_df['مؤشر SPI'].round(3)
        display_df['نسبة الإنجاز (%)'] = display_df['نسبة الإنجاز (%)'].round(1)
        display_df['الميزانية الإجمالية'] = display_df['الميزانية الإجمالية'].apply(lambda x: f"{x:,.0f} ريال")
        
        st.dataframe(display_df, use_container_width=True)
        
    except Exception as e:
        print(f"Error creating KPI dashboard: {e}")
        st.error("خطأ في إنشاء لوحة المراقبة")

def create_project_progress_chart(data_manager: DataManager, project_name: str) -> Optional[go.Figure]:
    """Create project progress chart over time"""
    try:
        progress_data = data_manager.get_progress_data(project_name)
        
        if progress_data.empty:
            return None
        
        # Sort by date
        progress_data['entry_date'] = pd.to_datetime(progress_data['entry_date'])
        progress_data = progress_data.sort_values('entry_date')
        
        fig = go.Figure()
        
        # Add planned completion line
        fig.add_trace(go.Scatter(
            x=progress_data['entry_date'],
            y=progress_data['planned_completion'],
            mode='lines+markers',
            name='نسبة الإنجاز المخطط',
            line=dict(color='blue', width=3),
            hovertemplate='<b>مخطط</b><br>' +
                         'التاريخ: %{x}<br>' +
                         'نسبة الإنجاز: %{y:.1f}%<extra></extra>'
        ))
        
        # Add actual completion line
        fig.add_trace(go.Scatter(
            x=progress_data['entry_date'],
            y=progress_data['actual_completion'],
            mode='lines+markers',
            name='نسبة الإنجاز الفعلي',
            line=dict(color='red', width=3),
            hovertemplate='<b>فعلي</b><br>' +
                         'التاريخ: %{x}<br>' +
                         'نسبة الإنجاز: %{y:.1f}%<extra></extra>'
        ))
        
        # Update layout
        fig.update_layout(
            title={
                'text': f'تقدم المشروع عبر الزمن - {project_name}',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 16}
            },
            xaxis_title='التاريخ',
            yaxis_title='نسبة الإنجاز (%)',
            yaxis=dict(range=[0, 100]),
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            font=dict(family="Arial", size=12),
            height=400
        )
        
        return fig
    except Exception as e:
        print(f"Error creating progress chart: {e}")
        return None

def create_cost_variance_chart(data_manager: DataManager, project_names: List[str]) -> Optional[go.Figure]:
    """Create cost variance comparison chart"""
    try:
        evm_calculator = EVMCalculator(data_manager)
        
        project_data = []
        for project_name in project_names:
            kpi = evm_calculator.calculate_project_kpi(project_name)
            if kpi:
                project_data.append({
                    'project_name': project_name,
                    'cost_variance': kpi['cv'],
                    'cost_variance_percent': kpi['cost_variance_percent']
                })
        
        if not project_data:
            return None
        
        df = pd.DataFrame(project_data)
        
        # Create bar chart
        fig = px.bar(
            df,
            x='project_name',
            y='cost_variance',
            title='انحراف التكلفة للمشاريع',
            labels={
                'project_name': 'اسم المشروع',
                'cost_variance': 'انحراف التكلفة (ريال)'
            },
            color='cost_variance',
            color_continuous_scale=['red', 'white', 'green']
        )
        
        # Add zero line
        fig.add_hline(y=0, line_dash="dash", line_color="black", annotation_text="لا يوجد انحراف")
        
        # Update layout
        fig.update_layout(
            font=dict(family="Arial", size=12),
            height=400,
            xaxis_tickangle=-45
        )
        
        return fig
    except Exception as e:
        print(f"Error creating cost variance chart: {e}")
        return None

def create_portfolio_overview_chart(data_manager: DataManager) -> Optional[go.Figure]:
    """Create portfolio overview chart with multiple metrics"""
    try:
        evm_calculator = EVMCalculator(data_manager)
        projects = data_manager.get_all_projects()
        
        if not projects:
            return None
        
        project_data = []
        for project in projects:
            kpi = evm_calculator.calculate_project_kpi(project['project_name'])
            if kpi:
                project_data.append(kpi)
        
        if not project_data:
            return None
        
        df = pd.DataFrame(project_data)
        
        # Create subplots
        from plotly.subplots import make_subplots
        
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=('مؤشر أداء التكلفة (CPI)', 'مؤشر أداء الجدولة (SPI)', 
                           'نسبة الإنجاز', 'انحراف التكلفة'),
            specs=[[{"type": "bar"}, {"type": "bar"}],
                   [{"type": "bar"}, {"type": "bar"}]]
        )
        
        # CPI chart
        fig.add_trace(
            go.Bar(x=df['project_name'], y=df['cpi'], name='CPI', 
                   marker_color='lightblue'),
            row=1, col=1
        )
        
        # SPI chart
        fig.add_trace(
            go.Bar(x=df['project_name'], y=df['spi'], name='SPI', 
                   marker_color='lightgreen'),
            row=1, col=2
        )
        
        # Completion chart
        fig.add_trace(
            go.Bar(x=df['project_name'], y=df['actual_completion'], name='نسبة الإنجاز', 
                   marker_color='orange'),
            row=2, col=1
        )
        
        # Cost variance chart
        fig.add_trace(
            go.Bar(x=df['project_name'], y=df['cv'], name='انحراف التكلفة', 
                   marker_color='red'),
            row=2, col=2
        )
        
        # Add reference lines
        fig.add_hline(y=1, line_dash="dash", line_color="gray", row=1, col=1)
        fig.add_hline(y=1, line_dash="dash", line_color="gray", row=1, col=2)
        fig.add_hline(y=0, line_dash="dash", line_color="gray", row=2, col=2)
        
        # Update layout
        fig.update_layout(
            title={
                'text': 'نظرة عامة على محفظة المشاريع',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 16}
            },
            font=dict(family="Arial", size=10),
            height=600,
            showlegend=False
        )
        
        # Update x-axis labels to be rotated
        fig.update_xaxes(tickangle=-45)
        
        return fig
    except Exception as e:
        print(f"Error creating portfolio overview chart: {e}")
        return None

def create_pie_charts(data_manager, project_names):
    """Create pie charts for project completion"""
    if not project_names:
        return None
    
    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
        
        # Create subplots for multiple pie charts
        cols = min(3, len(project_names))
        rows = (len(project_names) + cols - 1) // cols
        
        fig = make_subplots(
            rows=rows, cols=cols,
            specs=[[{"type": "pie"} for _ in range(cols)] for _ in range(rows)],
            subplot_titles=project_names
        )
        
        for i, project_name in enumerate(project_names):
            progress_data = data_manager.get_progress_data(project_name)
            
            if not progress_data.empty:
                latest = progress_data.iloc[-1]
                actual_completion = latest.get('actual_completion', 0)
                remaining = 100 - actual_completion
                
                row = (i // cols) + 1
                col = (i % cols) + 1
                
                fig.add_trace(
                    go.Pie(
                        labels=['Completed', 'Remaining'],
                        values=[actual_completion, remaining],
                        name=project_name,
                        marker_colors=['#2ecc71', '#ecf0f1'],
                        showlegend=True if i == 0 else False
                    ),
                    row=row, col=col
                )
        
        fig.update_layout(
            height=400 * rows,
            title_text="Project Completion Status",
            title_x=0.5
        )
        
        return fig
        
    except Exception as e:
        print(f"Error creating pie charts: {e}")
        return None


def create_bar_charts(data_manager, project_names):
    """Create bar charts for project comparison"""
    if not project_names:
        return None
    
    try:
        import plotly.graph_objects as go
        
        projects_info = []
        for project_name in project_names:
            project = data_manager.get_project_by_name(project_name)
            progress_data = data_manager.get_progress_data(project_name)
            
            if project and not progress_data.empty:
                latest = progress_data.iloc[-1]
                projects_info.append({
                    'name': project_name,
                    'budget': project.get('total_budget', 0),
                    'actual_cost': latest.get('actual_cost', 0),
                    'planned_completion': latest.get('planned_completion', 0),
                    'actual_completion': latest.get('actual_completion', 0)
                })
        
        if not projects_info:
            return None
        
        fig = go.Figure()
        
        # Budget vs Actual Cost
        fig.add_trace(go.Bar(
            name='Budget',
            x=[p['name'] for p in projects_info],
            y=[p['budget'] for p in projects_info],
            marker_color='#3498db'
        ))
        
        fig.add_trace(go.Bar(
            name='Actual Cost',
            x=[p['name'] for p in projects_info],
            y=[p['actual_cost'] for p in projects_info],
            marker_color='#e74c3c'
        ))
        
        fig.update_layout(
            title='Budget vs Actual Cost Comparison',
            xaxis_title='Projects',
            yaxis_title='Amount (SAR)',
            barmode='group',
            height=500
        )
        
        return fig
        
    except Exception as e:
        print(f"Error creating bar charts: {e}")
        return None


def create_gantt_chart(data_manager, project_names):
    """Create Gantt chart for project timeline"""
    if not project_names:
        return None
    
    try:
        import plotly.figure_factory as ff
        import pandas as pd
        from datetime import datetime
        
        gantt_data = []
        colors = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6']
        
        for i, project_name in enumerate(project_names):
            project = data_manager.get_project_by_name(project_name)
            
            if project:
                start_date = pd.to_datetime(project.get('start_date'))
                end_date = pd.to_datetime(project.get('end_date'))
                
                gantt_data.append(dict(
                    Task=project_name,
                    Start=start_date,
                    Finish=end_date,
                    Resource=f'Project {i+1}'
                ))
        
        if not gantt_data:
            return None
        
        fig = ff.create_gantt(
            gantt_data,
            colors=colors[:len(gantt_data)],
            index_col='Resource',
            title='Project Timeline',
            show_colorbar=True,
            bar_width=0.5,
            showgrid_x=True,
            showgrid_y=True
        )
        
        fig.update_layout(height=400)
        
        return fig
        
    except Exception as e:
        print(f"Error creating Gantt chart: {e}")
        return None


def create_dashboard_chart(data_manager, project_names):
    """Create comprehensive dashboard chart"""
    if not project_names:
        return None
    
    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
        
        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=('Progress Overview', 'Cost Comparison', 'Schedule Status', 'Risk Assessment'),
            specs=[[{"type": "bar"}, {"type": "scatter"}],
                   [{"type": "bar"}, {"type": "pie"}]]
        )
        
        project_data = []
        for project_name in project_names:
            project = data_manager.get_project_by_name(project_name)
            progress_data = data_manager.get_progress_data(project_name)
            
            if project and not progress_data.empty:
                latest = progress_data.iloc[-1]
                project_data.append({
                    'name': project_name,
                    'actual_completion': latest.get('actual_completion', 0),
                    'planned_completion': latest.get('planned_completion', 0),
                    'actual_cost': latest.get('actual_cost', 0),
                    'budget': project.get('total_budget', 0)
                })
        
        if not project_data:
            return None
        
        # Progress Overview
        fig.add_trace(
            go.Bar(
                x=[p['name'] for p in project_data],
                y=[p['actual_completion'] for p in project_data],
                name='Actual Progress',
                marker_color='#2ecc71',
                showlegend=False
            ),
            row=1, col=1
        )
        
        # Cost Comparison
        fig.add_trace(
            go.Scatter(
                x=[p['budget'] for p in project_data],
                y=[p['actual_cost'] for p in project_data],
                mode='markers',
                marker=dict(size=12, color='#3498db'),
                text=[p['name'] for p in project_data],
                showlegend=False
            ),
            row=1, col=2
        )
        
        # Schedule Status
        schedule_status = []
        for p in project_data:
            if p['actual_completion'] > p['planned_completion']:
                schedule_status.append('Ahead')
            elif p['actual_completion'] < p['planned_completion'] - 5:
                schedule_status.append('Behind')
            else:
                schedule_status.append('On Track')
        
        status_counts = {status: schedule_status.count(status) for status in set(schedule_status)}
        
        fig.add_trace(
            go.Bar(
                x=list(status_counts.keys()),
                y=list(status_counts.values()),
                marker_color=['#2ecc71', '#f39c12', '#e74c3c'],
                showlegend=False
            ),
            row=2, col=1
        )
        
        # Risk Assessment Pie
        risk_levels = []
        for p in project_data:
            cost_overrun = (p['actual_cost'] / p['budget'] - 1) * 100 if p['budget'] > 0 else 0
            schedule_delay = p['planned_completion'] - p['actual_completion']
            
            if cost_overrun > 10 or schedule_delay > 10:
                risk_levels.append('High Risk')
            elif cost_overrun > 5 or schedule_delay > 5:
                risk_levels.append('Medium Risk')
            else:
                risk_levels.append('Low Risk')
        
        risk_counts = {risk: risk_levels.count(risk) for risk in set(risk_levels)}
        
        fig.add_trace(
            go.Pie(
                labels=list(risk_counts.keys()),
                values=list(risk_counts.values()),
                marker_colors=['#2ecc71', '#f39c12', '#e74c3c'],
                showlegend=False
            ),
            row=2, col=2
        )
        
        fig.update_layout(
            height=800,
            title_text="Project Dashboard Overview",
            title_x=0.5
        )
        
        return fig
        
    except Exception as e:
        print(f"Error creating dashboard chart: {e}")
        return None
