import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import io
import re
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score
import warnings
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="Live Stream Analytics Dashboard Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Column mapping for cleaner names
COLUMN_MAPPING = {
    'ID Kreator': 'ID_Kreator',
    'Kreator': 'Kreator',
    'Nama panggilan': 'Nama_panggilan',
    'Waktu Live': 'Waktu_Live',
    'Durasi': 'Durasi',
    'Nilai barang dagangan bruto (LIVE) (Rp)': 'GMV_Bruto',
    'Produk yang ditambahkan': 'Produk_Added',
    'Produk Terjual': 'Produk_Terjual',
    'Pesanan SKU yang dibuat': 'Pesanan_SKU_Created',
    'Pesanan SKU (LIVE)': 'Pesanan_SKU_Live',
    'Produk yang terjual dari LIVE': 'Produk_Sold_Live',
    'Pembeli': 'Pembeli',
    'Harga Rata-Rata (Rp)': 'Harga_Rata_Rata',
    'Rasio pesanan per klik (LIVE)': 'Conversion_Rate',
    'GMV yang didapat dari LIVE (Rp)': 'GMV_Live',
    'Penonton': 'Penonton_Live_Stream',
    'Dilihat': 'Dilihat',
    'Durasi menonton rata-rata (Siaran LIVE)': 'Avg_Watch_Time',
    'Komentar': 'Komentar_Live',
    'Live Dibagikan': 'Dibagikan',
    'Suka pada LIVE': 'Suka_Live',
    'Pengikut baru (Video kreator)': 'New_Followers',
    'Produk Dilihat': 'Produk_Dilihat',
    'Klik Produk': 'Klik_Produk',
    'CTR': 'CTR'
}

# Required columns for validation
REQUIRED_COLUMNS = ['Kreator', 'GMV_Live', 'Penonton_Live_Stream', 'Pesanan_SKU_Live']

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #667eea;
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        transition: transform 0.2s;
        color: #333;
    }
    .metric-card:hover {
        transform: translateY(-2px);
    }
    .creator-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        border: 1px solid #e9ecef;
        margin-bottom: 1rem;
        transition: transform 0.2s;
        color: #333;
    }
    .creator-card:hover {
        transform: translateY(-2px);
    }
    .insight-card {
        background: linear-gradient(135deg, #e3f2fd 0%, #f1f8e9 100%);
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #4caf50;
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        color: #333;    
    }
    .warning-card {
        background: linear-gradient(135deg, #fff3e0 0%, #ffebee 100%);
        padding: 1.5rem;
        border-radius: 12px;
        border-left: 5px solid #ff9800;
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .stSelectbox > div > div {
        background-color: #f8f9fa;
        border-radius: 8px;
    }
    .performance-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: bold;
        margin: 0.2rem;
    }
    .badge-excellent {
        background: linear-gradient(135deg, #4caf50, #8bc34a);
        color: white;
    }
    .badge-good {
        background: linear-gradient(135deg, #2196f3, #03a9f4);
        color: white;
    }
    .badge-average {
        background: linear-gradient(135deg, #ff9800, #ffc107);
        color: white;
    }
    .badge-poor {
        background: linear-gradient(135deg, #f44336, #e91e63);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

def safe_numeric_conversion(value):
    """Safely convert value to numeric, handling various formats"""
    if pd.isna(value) or value == '' or value == '-':
        return 0
    
    str_value = str(value).strip()
    str_value = re.sub(r'[^\d.,\-]', '', str_value)
    
    if not str_value or str_value == '-':
        return 0
    
    str_value = str_value.replace(',', '.')
    
    try:
        return float(str_value)
    except (ValueError, TypeError):
        return 0

def parse_duration(duration_str):
    """Parse duration string to minutes"""
    if pd.isna(duration_str) or duration_str == '' or duration_str == '-':
        return 0
    
    duration_str = str(duration_str).strip()
    
    hours_match = re.search(r'(\d+)\s*h', duration_str, re.IGNORECASE)
    minutes_match = re.search(r'(\d+)\s*m', duration_str, re.IGNORECASE)
    
    hours = int(hours_match.group(1)) if hours_match else 0
    minutes = int(minutes_match.group(1)) if minutes_match else 0
    
    if hours == 0 and minutes == 0:
        numbers = re.findall(r'\d+', duration_str)
        if numbers:
            minutes = int(numbers[0])
            if len(numbers) > 1:
                hours = int(numbers[0])
                minutes = int(numbers[1])
    
    return hours * 60 + minutes

def parse_percentage(value):
    """Parse percentage string to float"""
    if pd.isna(value) or value == '' or value == '-':
        return 0.0
    
    str_value = str(value).strip()
    str_value = str_value.replace('%', '').replace(',', '.')
    
    try:
        return float(str_value)
    except (ValueError, TypeError):
        return 0.0

def calculate_performance_scores(df):
    """Calculate comprehensive performance scores"""
    if len(df) == 0:
        return df
    
    # Normalize metrics for scoring
    scaler = StandardScaler()
    
    # Key metrics for scoring
    scoring_metrics = ['GMV_Live', 'Penonton_Live_Stream', 'Pesanan_SKU_Live', 
                      'Engagement_Rate', 'Revenue_Per_Viewer', 'Conversion_Rate_Calc']
    
    # Handle missing metrics
    available_metrics = [col for col in scoring_metrics if col in df.columns]
    
    if not available_metrics:
        df['Performance_Score'] = 0
        return df
    
    # Create scoring matrix
    scoring_data = df[available_metrics].fillna(0)
    
    # Handle cases where all values are the same
    if scoring_data.nunique().sum() == 0:
        df['Performance_Score'] = 50
        return df
    
    # Normalize and calculate weighted score
    try:
        normalized = scaler.fit_transform(scoring_data)
        weights = [0.3, 0.2, 0.2, 0.1, 0.1, 0.1][:len(available_metrics)]
        
        # Calculate weighted score
        performance_scores = np.dot(normalized, weights)
        
        # Scale to 0-100
        min_score = performance_scores.min()
        max_score = performance_scores.max()
        
        if max_score != min_score:
            performance_scores = 100 * (performance_scores - min_score) / (max_score - min_score)
        else:
            performance_scores = np.full(len(performance_scores), 50)
            
        df['Performance_Score'] = performance_scores
        
    except Exception as e:
        df['Performance_Score'] = 50
    
    return df

def perform_creator_clustering(df):
    """Perform clustering analysis on creators"""
    if len(df) < 3:
        return df, None
    
    # Aggregate by creator
    creator_stats = df.groupby('Kreator').agg({
        'GMV_Live': 'mean',
        'Penonton_Live_Stream': 'mean',
        'Engagement_Rate': 'mean',
        'Revenue_Per_Viewer': 'mean',
        'Conversion_Rate_Calc': 'mean'
    }).fillna(0)
    
    if len(creator_stats) < 3:
        return df, None
    
    # Prepare data for clustering
    scaler = StandardScaler()
    
    try:
        scaled_data = scaler.fit_transform(creator_stats)
        
        # Determine optimal number of clusters
        n_clusters = min(4, len(creator_stats))
        
        # Perform clustering
        kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
        cluster_labels = kmeans.fit_predict(scaled_data)
        
        # Add cluster labels to creator stats
        creator_stats['Cluster'] = cluster_labels
        
        # Define cluster names
        cluster_names = {
            0: 'Rising Stars',
            1: 'Power Players',
            2: 'Consistent Performers',
            3: 'Niche Specialists'
        }
        
        creator_stats['Cluster_Name'] = creator_stats['Cluster'].map(
            lambda x: cluster_names.get(x, f'Group {x+1}')
        )
        
        # Merge back to main dataframe
        df = df.merge(creator_stats[['Cluster', 'Cluster_Name']], 
                     left_on='Kreator', right_index=True, how='left')
        
        return df, creator_stats
        
    except Exception as e:
        return df, None

def generate_insights(df):
    """Generate automated insights from data"""
    insights = []
    
    if len(df) == 0:
        return insights
    
    # Revenue insights
    top_revenue_creator = df.loc[df['GMV_Live'].idxmax(), 'Kreator']
    top_revenue = df['GMV_Live'].max()
    avg_revenue = df['GMV_Live'].mean()
    
    insights.append({
        'type': 'revenue',
        'title': 'Top Revenue Performer',
        'message': f"{top_revenue_creator} generated the highest revenue of {safe_format_currency(top_revenue)}, which is {((top_revenue/avg_revenue-1)*100):.1f}% above average.",
        'icon': '💰'
    })
    
    # Engagement insights
    if 'Engagement_Rate' in df.columns:
        high_engagement = df[df['Engagement_Rate'] > df['Engagement_Rate'].quantile(0.75)]
        if len(high_engagement) > 0:
            insights.append({
                'type': 'engagement',
                'title': 'High Engagement Creators',
                'message': f"{len(high_engagement)} creators have above-average engagement rates, with {high_engagement.loc[high_engagement['Engagement_Rate'].idxmax(), 'Kreator']} leading at {high_engagement['Engagement_Rate'].max():.1f}%.",
                'icon': '🚀'
            })
    
    # Conversion insights
    if 'Conversion_Rate_Calc' in df.columns:
        avg_conversion = df['Conversion_Rate_Calc'].mean()
        top_conversion_creator = df.loc[df['Conversion_Rate_Calc'].idxmax(), 'Kreator']
        top_conversion_rate = df['Conversion_Rate_Calc'].max()
        
        insights.append({
            'type': 'conversion',
            'title': 'Conversion Champion',
            'message': f"{top_conversion_creator} has the highest conversion rate at {top_conversion_rate:.2f}%, significantly outperforming the average of {avg_conversion:.2f}%.",
            'icon': '🎯'
        })
    
    # Duration insights
    if 'Durasi_Minutes' in df.columns:
        optimal_duration = df.loc[df['Revenue_Per_Viewer'].idxmax(), 'Durasi_Minutes']
        insights.append({
            'type': 'duration',
            'title': 'Optimal Stream Duration',
            'message': f"The most revenue-efficient stream duration appears to be around {optimal_duration:.0f} minutes based on revenue per viewer analysis.",
            'icon': '⏱️'
        })
    
    # Performance correlation
    if len(df) > 10:
        correlation = df[['Penonton_Live_Stream', 'GMV_Live']].corr().iloc[0, 1]
        if correlation > 0.7:
            insights.append({
                'type': 'correlation',
                'title': 'Strong Viewer-Revenue Correlation',
                'message': f"There's a strong positive correlation ({correlation:.2f}) between viewer count and revenue, suggesting effective monetization strategies.",
                'icon': '📈'
            })
    
    return insights

def create_advanced_charts(df):
    """Create advanced analytical charts"""
    charts = {}
    
    if len(df) == 0:
        return charts
    
    # 1. Performance correlation matrix
    numeric_cols = ['GMV_Live', 'Penonton_Live_Stream', 'Engagement_Rate', 
                   'Revenue_Per_Viewer', 'Conversion_Rate_Calc', 'Performance_Score']
    available_cols = [col for col in numeric_cols if col in df.columns]
    
    if len(available_cols) > 2:
        corr_matrix = df[available_cols].corr()
        
        charts['correlation'] = px.imshow(
            corr_matrix,
            title="Performance Metrics Correlation Matrix",
            color_continuous_scale='RdYlBu',
            aspect='auto'
        )
    
    # 2. Time series analysis (if date data available)
    if 'Waktu_Live' in df.columns and not df['Waktu_Live'].isna().all():
        daily_stats = df.groupby(df['Waktu_Live'].dt.date).agg({
            'GMV_Live': 'sum',
            'Penonton_Live_Stream': 'sum',
            'Pesanan_SKU_Live': 'sum'
        }).reset_index()
        
        charts['timeseries'] = px.line(
            daily_stats,
            x='Waktu_Live',
            y=['GMV_Live', 'Penonton_Live_Stream'],
            title="Daily Performance Trends",
            labels={'value': 'Metric Value', 'variable': 'Metrics'}
        )
    
    # 3. Performance distribution
    if 'Performance_Score' in df.columns:
        charts['performance_dist'] = px.histogram(
            df,
            x='Performance_Score',
            nbins=20,
            title="Performance Score Distribution",
            color_discrete_sequence=['#667eea']
        )
    
    return charts

def load_data(uploaded_file):
    """Load and clean the uploaded Excel file"""
    try:
        df = pd.read_excel(uploaded_file, skiprows=2)
        df.columns = df.columns.str.strip()
        df = df.rename(columns=COLUMN_MAPPING)
        
        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_columns:
            st.error(f"❌ Missing required columns: {missing_columns}")
            return None
        
        df = df.dropna(how='all')
        df = df.dropna(subset=['Kreator'])
        
        if 'Waktu_Live' in df.columns:
            df['Waktu_Live'] = pd.to_datetime(df['Waktu_Live'], errors='coerce')
        
        numeric_columns = [
            'GMV_Bruto', 'Produk_Added', 'Produk_Terjual', 'Pesanan_SKU_Created',
            'Pesanan_SKU_Live', 'Produk_Sold_Live', 'Pembeli', 'Harga_Rata_Rata',
            'GMV_Live', 'Penonton_Live_Stream', 'Dilihat', 'Avg_Watch_Time',
            'Komentar_Live', 'Dibagikan', 'Suka_Live', 'New_Followers',
            'Produk_Dilihat', 'Klik_Produk'
        ]
        
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].apply(safe_numeric_conversion)
        
        if 'Durasi' in df.columns:
            df['Durasi_Minutes'] = df['Durasi'].apply(parse_duration)
        
        percentage_columns = ['Conversion_Rate', 'CTR']
        for col in percentage_columns:
            if col in df.columns:
                df[col] = df[col].apply(parse_percentage)
        
        # Calculate derived metrics
        df['Engagement_Rate'] = 0
        if all(col in df.columns for col in ['Suka_Live', 'Komentar_Live', 'Dibagikan', 'Penonton_Live_Stream']):
            df['Engagement_Rate'] = np.where(
                df['Penonton_Live_Stream'] > 0,
                ((df['Suka_Live'] + df['Komentar_Live'] + df['Dibagikan']) / df['Penonton_Live_Stream']) * 100,
                0
            )
        
        df['Revenue_Per_Viewer'] = np.where(
            df['Penonton_Live_Stream'] > 0,
            df['GMV_Live'] / df['Penonton_Live_Stream'],
            0
        )
        
        df['Conversion_Rate_Calc'] = np.where(
            df['Penonton_Live_Stream'] > 0,
            (df['Pesanan_SKU_Live'] / df['Penonton_Live_Stream']) * 100,
            0
        )
        
        # Calculate performance scores
        df = calculate_performance_scores(df)
        
        # Perform clustering
        df, cluster_stats = perform_creator_clustering(df)
        
        df = df.fillna(0)
        
        return df
    
    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        return None

def safe_format_number(num):
    """Safely format numbers"""
    try:
        if pd.isna(num) or num == 0:
            return "0"
        if abs(num) >= 1000000:
            return f"{num/1000000:.1f}M"
        elif abs(num) >= 1000:
            return f"{num/1000:.1f}K"
        else:
            return f"{num:,.0f}"
    except:
        return "0"

def safe_format_currency(num):
    """Safely format currency"""
    try:
        if pd.isna(num) or num == 0:
            return "Rp 0"
        return f"Rp {num:,.0f}"
    except:
        return "Rp 0"

def create_safe_chart(df, chart_type, **kwargs):
    """Create charts with error handling"""
    try:
        if chart_type == 'bar':
            return px.bar(df, **kwargs)
        elif chart_type == 'scatter':
            return px.scatter(df, **kwargs)
        elif chart_type == 'histogram':
            return px.histogram(df, **kwargs)
        elif chart_type == 'line':
            return px.line(df, **kwargs)
        else:
            return go.Figure()
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return go.Figure()

def get_performance_badge(score):
    """Get performance badge based on score"""
    if score >= 80:
        return '<span class="performance-badge badge-excellent">🏆 Excellent</span>'
    elif score >= 60:
        return '<span class="performance-badge badge-good">⭐ Good</span>'
    elif score >= 40:
        return '<span class="performance-badge badge-average">📊 Average</span>'
    else:
        return '<span class="performance-badge badge-poor">📉 Needs Improvement</span>'

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1 style="color: white; margin: 0; text-align: center;">
            📊 Live Stream Analytics Dashboard Pro
        </h1>
        <p style="color: white; margin: 0; text-align: center; opacity: 0.9;">
            Advanced Analytics & AI-Powered Insights for Live Stream Performance
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.title("📁 Data Upload")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xls'],
        help="Upload your daily live stream data file"
    )
    
    if uploaded_file is not None:
        with st.spinner("🔄 Processing data with AI analytics..."):
            df = load_data(uploaded_file)
        
        if df is not None and len(df) > 0:
            st.success(f"✅ Data loaded successfully! {len(df)} records processed with AI insights.")
            
            # Generate insights
            insights = generate_insights(df)
            
            # Data quality check
            quality_issues = []
            if df['GMV_Live'].sum() == 0:
                quality_issues.append("No GMV data found")
            if df['Penonton_Live_Stream'].sum() == 0:
                quality_issues.append("No viewer data found")
            if len(df['Kreator'].unique()) == 1:
                quality_issues.append("Only one creator found")
            
            if quality_issues:
                st.warning(f"⚠️ Data quality issues detected: {', '.join(quality_issues)}")
            
            # Enhanced sidebar info
            st.sidebar.subheader("📊 Data Overview")
            st.sidebar.info(f"""
            **Total Creators:** {len(df['Kreator'].unique())}
            **Total Records:** {len(df)}
            **Date Range:** {df['Waktu_Live'].min().strftime('%Y-%m-%d') if 'Waktu_Live' in df.columns and not df['Waktu_Live'].isna().all() else 'Not available'}
            **Avg Performance Score:** {df['Performance_Score'].mean():.1f}/100
            """)
            
            # Advanced filters
            st.sidebar.subheader("🔍 Advanced Filters")
            
            # Creator filter
            all_creators = df['Kreator'].unique()
            default_creators = all_creators[:5] if len(all_creators) > 5 else all_creators
            
            selected_creators = st.sidebar.multiselect(
                "Select Creators:",
                options=all_creators,
                default=default_creators
            )
            
            # Performance filter
            if 'Performance_Score' in df.columns:
                perf_range = st.sidebar.slider(
                    "Performance Score Range:",
                    min_value=float(df['Performance_Score'].min()),
                    max_value=float(df['Performance_Score'].max()),
                    value=(float(df['Performance_Score'].min()), float(df['Performance_Score'].max())),
                    step=1.0
                )
                
                df = df[(df['Performance_Score'] >= perf_range[0]) & (df['Performance_Score'] <= perf_range[1])]
            
            # Cluster filter
            if 'Cluster_Name' in df.columns:
                cluster_filter = st.sidebar.multiselect(
                    "Select Creator Clusters:",
                    options=df['Cluster_Name'].unique(),
                    default=df['Cluster_Name'].unique()
                )
                df = df[df['Cluster_Name'].isin(cluster_filter)]
            
            if selected_creators:
                df_filtered = df[df['Kreator'].isin(selected_creators)]
            else:
                df_filtered = df
            
            # Date filter
            if 'Waktu_Live' in df.columns and not df['Waktu_Live'].isna().all():
                date_range = st.sidebar.date_input(
                    "Select Date Range:",
                    value=(df['Waktu_Live'].min().date(), df['Waktu_Live'].max().date()),
                    min_value=df['Waktu_Live'].min().date(),
                    max_value=df['Waktu_Live'].max().date()
                )
                
                if len(date_range) == 2:
                    start_date, end_date = date_range
                    df_filtered = df_filtered[
                        (df_filtered['Waktu_Live'].dt.date >= start_date) & 
                        (df_filtered['Waktu_Live'].dt.date <= end_date)
                    ]
            
            # Main dashboard tabs
            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "🎯 Smart Insights", "📊 Overview", "💰 Revenue Analytics", 
                "👥 Audience Analytics", "🛍️ Product Analytics", "🏆 Performance Rankings",
                "🔬 Advanced Analytics"
            ])
            
            with tab1:
                st.header("🎯 AI-Powered Smart Insights")
                
                # Display insights
                if insights:
                    col1, col2 = st.columns(2)
                    
                    for i, insight in enumerate(insights):
                        with col1 if i % 2 == 0 else col2:
                            st.markdown(f"""
                            <div class="insight-card">
                                <h4>{insight['icon']} {insight['title']}</h4>
                                <p>{insight['message']}</p>
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.info("🔄 Analyzing data to generate insights...")
                
                # Key performance indicators
                st.subheader("📈 Key Performance Indicators")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    avg_performance = df_filtered['Performance_Score'].mean()
                    st.metric(
                        label="Average Performance Score",
                        value=f"{avg_performance:.1f}/100",
                        delta=f"{len(df_filtered)} creators"
                    )
                
                with col2:
                    total_gmv = df_filtered['GMV_Live'].sum()
                    avg_gmv = df_filtered['GMV_Live'].mean()
                    st.metric(
                        label="Total GMV",
                        value=safe_format_currency(total_gmv),
                        delta=f"Avg: {safe_format_currency(avg_gmv)}"
                    )
                
                with col3:
                    total_engagement = df_filtered['Engagement_Rate'].mean()
                    top_engagement = df_filtered['Engagement_Rate'].max()
                    st.metric(
                        label="Engagement Rate",
                        value=f"{total_engagement:.1f}%",
                        delta=f"Peak: {top_engagement:.1f}%"
                    )
                
                with col4:
                    avg_conversion = df_filtered['Conversion_Rate_Calc'].mean()
                    top_conversion = df_filtered['Conversion_Rate_Calc'].max()
                    st.metric(
                        label="Conversion Rate",
                        value=f"{avg_conversion:.2f}%",
                        delta=f"Peak: {top_conversion:.2f}%"
                    )
                
                # Performance recommendations
                st.subheader("💡 Performance Recommendations")
                
                # Identify improvement areas
                low_performers = df_filtered[df_filtered['Performance_Score'] < 40]
                if len(low_performers) > 0:
                    st.markdown("""
                    <div class="warning-card">
                        <h4 style="color: #333;">⚠️ Attention Required</h4>
                        <p style="color: #333;" >Several creators have performance scores below 40. Consider:</p>
                        <ul style="color: #333;">
                            <li>Reviewing content strategy and engagement tactics</li>
                            <li>Analyzing successful creators' approaches</li>
                            <li>Implementing targeted improvement programs</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Top performers insights
                top_performers = df_filtered.nlargest(3, 'Performance_Score')
                if len(top_performers) > 0:
                    st.markdown("""
                    <div class="insight-card">
                        <h4>🌟 Top Performers Insights</h4>
                        <p>Learn from the best performing creators:</p>
                        <ul>
                            <li>Average engagement rate: {:.1f}%</li>
                            <li>Average revenue per viewer: {}</li>
                            <li>Key success factors: High engagement, consistent streaming</li>
                        </ul>
                    </div>
                    """.format(
                        top_performers['Engagement_Rate'].mean(),
                        safe_format_currency(top_performers['Revenue_Per_Viewer'].mean())
                    ), unsafe_allow_html=True)
            
            with tab2:
                st.header("📊 Performance Overview")
                
                # Summary metrics
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    total_revenue = df_filtered['GMV_Live'].sum()
                    st.markdown(f"""
                    <div class="metric-card">
                        <h3>💰 Total Revenue</h3>
                        <h2>{safe_format_currency(total_revenue)}</h2>
                        <p>From {len(df_filtered)} streams</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    total_viewers = df_filtered['Penonton_Live_Stream'].sum()
                    st.markdown(f"""
                    <div class="metric-card">
                        <h3>👥 Total Viewers</h3>
                        <h2>{safe_format_number(total_viewers)}</h2>
                        <p>Avg: {safe_format_number(df_filtered['Penonton_Live_Stream'].mean())}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    total_orders = df_filtered['Pesanan_SKU_Live'].sum()
                    st.markdown(f"""
                    <div class="metric-card">
                        <h3>🛍️ Total Orders</h3>
                        <h2>{safe_format_number(total_orders)}</h2>
                        <p>Avg: {safe_format_number(df_filtered['Pesanan_SKU_Live'].mean())}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    avg_engagement = df_filtered['Engagement_Rate'].mean()
                    st.markdown(f"""
                    <div class="metric-card">
                        <h3>🚀 Engagement Rate</h3>
                        <h2>{avg_engagement:.1f}%</h2>
                        <p>Peak: {df_filtered['Engagement_Rate'].max():.1f}%</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Charts
                col1, col2 = st.columns(2)
                
                with col1:
                    # Revenue by creator
                    creator_revenue = df_filtered.groupby('Kreator')['GMV_Live'].sum().sort_values(ascending=False).head(10)
                    fig_revenue = create_safe_chart(
                        creator_revenue.reset_index(),
                        'bar',
                        x='Kreator',
                        y='GMV_Live',
                        title="Top 10 Creators by Revenue",
                        color='GMV_Live',
                        color_continuous_scale='Blues'
                    )
                    fig_revenue.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_revenue, use_container_width=True)
                
                with col2:
                    # Viewers by creator
                    creator_viewers = df_filtered.groupby('Kreator')['Penonton_Live_Stream'].sum().sort_values(ascending=False).head(10)
                    fig_viewers = create_safe_chart(
                        creator_viewers.reset_index(),
                        'bar',
                        x='Kreator',
                        y='Penonton_Live_Stream',
                        title="Top 10 Creators by Viewers",
                        color='Penonton_Live_Stream',
                        color_continuous_scale='Greens'
                    )
                    fig_viewers.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_viewers, use_container_width=True)
                
                # Performance scatter plot
                st.subheader("📈 Performance Correlation Analysis")
                fig_scatter = create_safe_chart(
                    df_filtered,
                    'scatter',
                    x='Penonton_Live_Stream',
                    y='GMV_Live',
                    color='Performance_Score',
                    size='Engagement_Rate',
                    hover_data=['Kreator'],
                    title="Revenue vs Viewers (sized by Engagement, colored by Performance Score)",
                    color_continuous_scale='Viridis'
                )
                st.plotly_chart(fig_scatter, use_container_width=True)
            
            with tab3:
                st.header("💰 Revenue Analytics")
                
                # Revenue metrics
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    revenue_per_viewer = df_filtered['Revenue_Per_Viewer'].mean()
                    st.metric(
                        label="Revenue per Viewer",
                        value=safe_format_currency(revenue_per_viewer),
                        delta=f"Max: {safe_format_currency(df_filtered['Revenue_Per_Viewer'].max())}"
                    )
                
                with col2:
                    avg_order_value = df_filtered['GMV_Live'].sum() / df_filtered['Pesanan_SKU_Live'].sum() if df_filtered['Pesanan_SKU_Live'].sum() > 0 else 0
                    st.metric(
                        label="Average Order Value",
                        value=safe_format_currency(avg_order_value),
                        delta=f"Total Orders: {safe_format_number(df_filtered['Pesanan_SKU_Live'].sum())}"
                    )
                
                with col3:
                    conversion_rate = df_filtered['Conversion_Rate_Calc'].mean()
                    st.metric(
                        label="Conversion Rate",
                        value=f"{conversion_rate:.2f}%",
                        delta=f"Best: {df_filtered['Conversion_Rate_Calc'].max():.2f}%"
                    )
                
                # Revenue trends
                col1, col2 = st.columns(2)
                
                with col1:
                    # Revenue distribution
                    fig_revenue_dist = create_safe_chart(
                        df_filtered,
                        'histogram',
                        x='GMV_Live',
                        title="Revenue Distribution",
                        nbins=20,
                        color_discrete_sequence=['#667eea']
                    )
                    st.plotly_chart(fig_revenue_dist, use_container_width=True)
                
                with col2:
                    # Revenue per viewer by creator
                    creator_rpv = df_filtered.groupby('Kreator')['Revenue_Per_Viewer'].mean().sort_values(ascending=False).head(10)
                    fig_rpv = create_safe_chart(
                        creator_rpv.reset_index(),
                        'bar',
                        x='Kreator',
                        y='Revenue_Per_Viewer',
                        title="Top 10 Revenue per Viewer",
                        color='Revenue_Per_Viewer',
                        color_continuous_scale='Reds'
                    )
                    fig_rpv.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_rpv, use_container_width=True)
                
                # Revenue efficiency analysis
                st.subheader("📊 Revenue Efficiency Analysis")
                
                # Create efficiency metrics
                efficiency_data = df_filtered.groupby('Kreator').agg({
                    'GMV_Live': 'sum',
                    'Penonton_Live_Stream': 'sum',
                    'Revenue_Per_Viewer': 'mean',
                    'Conversion_Rate_Calc': 'mean'
                }).reset_index()
                
                fig_efficiency = create_safe_chart(
                    efficiency_data,
                    'scatter',
                    x='Conversion_Rate_Calc',
                    y='Revenue_Per_Viewer',
                    color='GMV_Live',
                    size='Penonton_Live_Stream',
                    hover_data=['Kreator'],
                    title="Revenue Efficiency: Conversion Rate vs Revenue per Viewer",
                    color_continuous_scale='Plasma'
                )
                st.plotly_chart(fig_efficiency, use_container_width=True)
            
            with tab4:
                st.header("👥 Audience Analytics")
                
                # Audience metrics
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    avg_viewers = df_filtered['Penonton_Live_Stream'].mean()
                    st.metric(
                        label="Average Viewers",
                        value=safe_format_number(avg_viewers),
                        delta=f"Peak: {safe_format_number(df_filtered['Penonton_Live_Stream'].max())}"
                    )
                
                with col2:
                    if 'Avg_Watch_Time' in df_filtered.columns:
                        avg_watch_time = df_filtered['Avg_Watch_Time'].mean()
                        st.metric(
                            label="Avg Watch Time",
                            value=f"{avg_watch_time:.1f} min",
                            delta=f"Max: {df_filtered['Avg_Watch_Time'].max():.1f} min"
                        )
                
                with col3:
                    if 'New_Followers' in df_filtered.columns:
                        total_new_followers = df_filtered['New_Followers'].sum()
                        st.metric(
                            label="New Followers",
                            value=safe_format_number(total_new_followers),
                            delta=f"Avg per stream: {safe_format_number(df_filtered['New_Followers'].mean())}"
                        )
                
                # Audience engagement charts
                col1, col2 = st.columns(2)
                
                with col1:
                    # Engagement rate distribution
                    fig_engagement = create_safe_chart(
                        df_filtered,
                        'histogram',
                        x='Engagement_Rate',
                        title="Engagement Rate Distribution",
                        nbins=20,
                        color_discrete_sequence=['#4caf50']
                    )
                    st.plotly_chart(fig_engagement, use_container_width=True)
                
                with col2:
                    # Viewer engagement correlation
                    fig_viewer_engagement = create_safe_chart(
                        df_filtered,
                        'scatter',
                        x='Penonton_Live_Stream',
                        y='Engagement_Rate',
                        color='GMV_Live',
                        hover_data=['Kreator'],
                        title="Viewers vs Engagement Rate",
                        color_continuous_scale='Turbo'
                    )
                    st.plotly_chart(fig_viewer_engagement, use_container_width=True)
                
                # Audience retention analysis
                if 'Durasi_Minutes' in df_filtered.columns and 'Avg_Watch_Time' in df_filtered.columns:
                    st.subheader("📺 Audience Retention Analysis")
                    
                    # Calculate retention rate
                    df_filtered['Retention_Rate'] = np.where(
                        df_filtered['Durasi_Minutes'] > 0,
                        (df_filtered['Avg_Watch_Time'] / df_filtered['Durasi_Minutes']) * 100,
                        0
                    )
                    
                    retention_data = df_filtered.groupby('Kreator')['Retention_Rate'].mean().sort_values(ascending=False).head(10)
                    fig_retention = create_safe_chart(
                        retention_data.reset_index(),
                        'bar',
                        x='Kreator',
                        y='Retention_Rate',
                        title="Top 10 Audience Retention Rates",
                        color='Retention_Rate',
                        color_continuous_scale='YlOrRd'
                    )
                    fig_retention.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_retention, use_container_width=True)
            
            with tab5:
                st.header("🛍️ Product Analytics")
                
                # Product metrics
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if 'Produk_Added' in df_filtered.columns:
                        total_products = df_filtered['Produk_Added'].sum()
                        st.metric(
                            label="Products Added",
                            value=safe_format_number(total_products),
                            delta=f"Avg: {safe_format_number(df_filtered['Produk_Added'].mean())}"
                        )
                
                with col2:
                    if 'Produk_Terjual' in df_filtered.columns:
                        total_sold = df_filtered['Produk_Terjual'].sum()
                        st.metric(
                            label="Products Sold",
                            value=safe_format_number(total_sold),
                            delta=f"Avg: {safe_format_number(df_filtered['Produk_Terjual'].mean())}"
                        )
                
                with col3:
                    if 'CTR' in df_filtered.columns:
                        avg_ctr = df_filtered['CTR'].mean()
                        st.metric(
                            label="Click-Through Rate",
                            value=f"{avg_ctr:.2f}%",
                            delta=f"Best: {df_filtered['CTR'].max():.2f}%"
                        )
                
                # Product performance charts
                if 'Produk_Added' in df_filtered.columns and 'Produk_Terjual' in df_filtered.columns:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Product conversion rate
                        df_filtered['Product_Conversion'] = np.where(
                            df_filtered['Produk_Added'] > 0,
                            (df_filtered['Produk_Terjual'] / df_filtered['Produk_Added']) * 100,
                            0
                        )
                        
                        fig_product_conv = create_safe_chart(
                            df_filtered,
                            'histogram',
                            x='Product_Conversion',
                            title="Product Conversion Rate Distribution",
                            nbins=20,
                            color_discrete_sequence=['#ff9800']
                        )
                        st.plotly_chart(fig_product_conv, use_container_width=True)
                    
                    with col2:
                        # Top performers by product metrics
                        product_performance = df_filtered.groupby('Kreator').agg({
                            'Produk_Added': 'sum',
                            'Produk_Terjual': 'sum',
                            'Product_Conversion': 'mean'
                        }).reset_index()
                        
                        fig_product_perf = create_safe_chart(
                            product_performance.head(10),
                            'scatter',
                            x='Produk_Added',
                            y='Produk_Terjual',
                            color='Product_Conversion',
                            size='Product_Conversion',
                            hover_data=['Kreator'],
                            title="Product Performance: Added vs Sold",
                            color_continuous_scale='Spectral'
                        )
                        st.plotly_chart(fig_product_perf, use_container_width=True)
            
            with tab6:
                st.header("🏆 Performance Rankings")
                
                # Performance ranking table
                st.subheader("📊 Creator Performance Leaderboard")
                
                # Calculate comprehensive rankings
                ranking_data = df_filtered.groupby('Kreator').agg({
                    'GMV_Live': 'sum',
                    'Penonton_Live_Stream': 'sum',
                    'Engagement_Rate': 'mean',
                    'Revenue_Per_Viewer': 'mean',
                    'Conversion_Rate_Calc': 'mean',
                    'Performance_Score': 'mean'
                }).reset_index()
                
                ranking_data = ranking_data.sort_values('Performance_Score', ascending=False)
                ranking_data['Rank'] = range(1, len(ranking_data) + 1)
                
                # Display top performers
                for i, row in ranking_data.head(10).iterrows():
                    performance_badge = get_performance_badge(row['Performance_Score'])
                    
                    st.markdown(f"""
                    <div class="creator-card">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <h3>#{row['Rank']} {row['Kreator']}</h3>
                                {performance_badge}
                            </div>
                            <div style="text-align: right;">
                                <h4>{safe_format_currency(row['GMV_Live'])}</h4>
                                <p>{safe_format_number(row['Penonton_Live_Stream'])} viewers</p>
                            </div>
                        </div>
                        <div style="margin-top: 1rem;">
                            <div style="display: flex; justify-content: space-between;">
                                <span>Engagement: {row['Engagement_Rate']:.1f}%</span>
                                <span>Revenue/Viewer: {safe_format_currency(row['Revenue_Per_Viewer'])}</span>
                                <span>Conversion: {row['Conversion_Rate_Calc']:.2f}%</span>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Performance comparison chart
                st.subheader("📈 Performance Comparison")
                
                fig_comparison = go.Figure()
                
                # Add traces for different metrics
                fig_comparison.add_trace(go.Scatter(
                    x=ranking_data['Kreator'][:10],
                    y=ranking_data['Performance_Score'][:10],
                    mode='lines+markers',
                    name='Performance Score',
                    line=dict(color='#667eea', width=3),
                    marker=dict(size=8)
                ))
                
                fig_comparison.update_layout(
                    title="Top 10 Creators Performance Comparison",
                    xaxis_title="Creator",
                    yaxis_title="Performance Score",
                    xaxis_tickangle=-45,
                    height=400
                )
                
                st.plotly_chart(fig_comparison, use_container_width=True)
            
            with tab7:
                st.header("🔬 Advanced Analytics")
                
                # Advanced charts
                advanced_charts = create_advanced_charts(df_filtered)
                
                if 'correlation' in advanced_charts:
                    st.subheader("📊 Correlation Matrix")
                    st.plotly_chart(advanced_charts['correlation'], use_container_width=True)
                
                if 'timeseries' in advanced_charts:
                    st.subheader("📈 Time Series Analysis")
                    st.plotly_chart(advanced_charts['timeseries'], use_container_width=True)
                
                if 'performance_dist' in advanced_charts:
                    st.subheader("📊 Performance Distribution")
                    st.plotly_chart(advanced_charts['performance_dist'], use_container_width=True)
                
                # Clustering analysis
                if 'Cluster_Name' in df_filtered.columns:
                    st.subheader("🎯 Creator Clustering Analysis")
                    
                    cluster_summary = df_filtered.groupby('Cluster_Name').agg({
                        'GMV_Live': 'mean',
                        'Penonton_Live_Stream': 'mean',
                        'Engagement_Rate': 'mean',
                        'Performance_Score': 'mean'
                    }).reset_index()
                    
                    fig_clusters = create_safe_chart(
                        cluster_summary,
                        'bar',
                        x='Cluster_Name',
                        y='Performance_Score',
                        color='Cluster_Name',
                        title="Average Performance Score by Cluster",
                        color_discrete_sequence=['#667eea', '#764ba2', '#f093fb', '#4caf50']
                    )
                    st.plotly_chart(fig_clusters, use_container_width=True)
                
                # Predictive analytics
                st.subheader("🔮 Predictive Analytics")
                
                try:
                    # Simple linear regression for revenue prediction
                    X = df_filtered[['Penonton_Live_Stream', 'Engagement_Rate']].fillna(0)
                    y = df_filtered['GMV_Live']
                    
                    if len(X) > 10:
                        model = LinearRegression()
                        model.fit(X, y)
                        
                        predictions = model.predict(X)
                        r2 = r2_score(y, predictions)
                        
                        st.info(f"Revenue Prediction Model R² Score: {r2:.3f}")
                        
                        # Create prediction vs actual chart
                        fig_prediction = go.Figure()
                        fig_prediction.add_trace(go.Scatter(
                            x=y,
                            y=predictions,
                            mode='markers',
                            name='Predictions',
                            marker=dict(color='#667eea', size=8)
                        ))
                        
                        # Add perfect prediction line
                        min_val = min(y.min(), predictions.min())
                        max_val = max(y.max(), predictions.max())
                        fig_prediction.add_trace(go.Scatter(
                            x=[min_val, max_val],
                            y=[min_val, max_val],
                            mode='lines',
                            name='Perfect Prediction',
                            line=dict(color='red', dash='dash')
                        ))
                        
                        fig_prediction.update_layout(
                            title="Revenue Prediction vs Actual",
                            xaxis_title="Actual Revenue",
                            yaxis_title="Predicted Revenue",
                            height=400
                        )
                        
                        st.plotly_chart(fig_prediction, use_container_width=True)
                        
                except Exception as e:
                    st.warning("Not enough data for predictive analysis")
                
                # Export functionality
                st.subheader("📥 Export Data")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if st.button("📊 Export Filtered Data"):
                        csv = df_filtered.to_csv(index=False)
                        st.download_button(
                            label="Download CSV",
                            data=csv,
                            file_name=f"live_stream_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv"
                        )
                
                with col2:
                    if st.button("📈 Export Performance Report"):
                        report_data = ranking_data.to_csv(index=False)
                        st.download_button(
                            label="Download Performance Report",
                            data=report_data,
                            file_name=f"performance_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv"
                        )
        
        else:
            st.error("❌ No valid data found in the uploaded file.")
    
    else:
        st.info("👆 Please upload an Excel file to get started with the analytics dashboard.")
        
        # Display sample data format
        st.subheader("📋 Expected Data Format")
        st.info("""
        Your Excel file should contain the following columns:
        - Kreator (Creator name)
        - GMV Live (Revenue from live stream)
        - Penonton Live Stream (Number of viewers)
        - Pesanan SKU Live (Orders from live stream)
        - And other performance metrics...
        """)

if __name__ == "__main__":
    main()
