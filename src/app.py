import pandas as pd
from dash import Dash, dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Load the dataset
file_path = "https://github.com/clarence-ck/NLB_Top100/raw/refs/heads/main/src/Top_100_OD_Titles_CY2020_to_2023.xlsx"
data = pd.read_excel(file_path, sheet_name='Sheet1')

# Convert 'Title Publication Date' to datetime
data['Title Publication Date'] = pd.to_datetime(data['Title Publication Date'], format='%Y-%m-%d')
data['Publication Year'] = data['Title Publication Date'].dt.year

# Convert other columns to appropriate data types
data['Item Media'] = data['Item Media'].str.title()
data['Title Native Name'] = data['Title Native Name'].astype(str)
data['Title Author'] = data['Title Author'].astype(str)
data['Title Publisher'] = data['Title Publisher'].astype(str)
data['Txn Calendar Year'] = data['Txn Calendar Year'].astype(int)
data['Subject'] = data['Subject'].astype(str)

# KPI Preprocessing
total_titles = data['Title Native Name'].nunique()
total_authors = data['Title Author'].nunique()
total_publishers = data['Title Publisher'].nunique()
earliest_publication = data['Title Publication Date'].min().strftime('%Y-%m-%d')
latest_publication = data['Title Publication Date'].max().strftime('%Y-%m-%d')

# Initialize Dash app with a Bootstrap theme
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server  # For deployment

# Color Palette
colors = {
    'background': '#FFFFFF',
    'text': '#2F3E46',
    'primary': '#305F72',
    'secondary': '#F28F3B',
    'accent': '#9C3D54',
}

# App Layout
app.layout = dbc.Container([
    # Navigation Bar
    dbc.NavbarSimple(
        brand="Top 100 OD Titles Dashboard (2020 - 2023)",
        brand_href="#",
        color=colors['primary'],
        dark=True,
        className="mb-4",
    ),

    # Filters
    dbc.Row([
        dbc.Col([
            html.Label('Select Year(s):', style={'font-weight': 'bold'}),
            dcc.Dropdown(
                id='year-filter',
                options=[{'label': str(year), 'value': year} for year in sorted(data['Txn Calendar Year'].unique())],
                value=sorted(data['Txn Calendar Year'].unique()),
                multi=True,
                placeholder="Select Year(s)",
            )
        ], width=3),
        dbc.Col([
            html.Label('Select Subject(s):', style={'font-weight': 'bold'}),
            dcc.Dropdown(
                id='subject-filter',
                options=[{'label': subj, 'value': subj} for subj in sorted(data['Subject'].unique())],
                value=[],
                multi=True,
                placeholder="Select Subject(s)",
            )
        ], width=3),
        dbc.Col([
            html.Label('Select Media Type(s):', style={'font-weight': 'bold'}),
            dcc.Dropdown(
                id='media-filter',
                options=[{'label': media, 'value': media} for media in sorted(data['Item Media'].unique())],
                value=[],
                multi=True,
                placeholder="Select Media Type(s)",
            )
        ], width=3),
        dbc.Col([
            html.Label('Select Publication Date Range:', style={'font-weight': 'bold'}),
            dcc.DatePickerRange(
                id='publication-date-filter',
                min_date_allowed=data['Title Publication Date'].min(),
                max_date_allowed=data['Title Publication Date'].max(),
                start_date=data['Title Publication Date'].min(),
                end_date=data['Title Publication Date'].max(),
                display_format='YYYY-MM-DD',
                style={'width': '100%'}
            )
        ], width=3),
    ], className="mb-4"),

    # KPI Cards
    dbc.Row([
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H4(id='total-titles', className="card-title text-center", style={'color': colors['accent']}),
                html.P("Unique Titles", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=2),
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H4(id='total-authors', className="card-title text-center", style={'color': colors['accent']}),
                html.P("Unique Authors", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=2),
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H4(id='total-publishers', className="card-title text-center", style={'color': colors['accent']}),
                html.P("Unique Publishers", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=2),
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H4(f"{earliest_publication}", className="card-title text-center", style={'color': colors['accent']}),
                html.P("Earliest Publication Date", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=3),
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H4(f"{latest_publication}", className="card-title text-center", style={'color': colors['accent']}),
                html.P("Latest Publication Date", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=3),
    ], className="mb-5"),

    # Tabs for organizing content
    dbc.Tabs([
        # Overview Tab
        dbc.Tab(label='Overview', tab_id='tab-overview', children=[
            dbc.Row([
                dbc.Col(dcc.Graph(id='media-type-distribution'), width=6),
                dbc.Col(dcc.Graph(id='top-subjects-bar'), width=6),
            ], className="mb-4"),

            dbc.Row([
                dbc.Col(dcc.Graph(id='publication-year-distribution'), width=6),
                dbc.Col(dcc.Graph(id='yearly-trend-chart'), width=6),
            ]),
        ]),

        # Detailed Analysis Tab
        dbc.Tab(label='Detailed Analysis', tab_id='tab-detailed', children=[
            dbc.Row([
                dbc.Col([
                    html.Label('Select Top N Authors:', style={'font-weight': 'bold'}),
                    dcc.Slider(
                        id='top-authors-slider',
                        min=5,
                        max=20,
                        step=1,
                        value=10,
                        marks={i: str(i) for i in range(5, 21, 5)},
                        tooltip={"placement": "bottom", "always_visible": True},
                    ),
                    dcc.Graph(id='author-heatmap')
                ], width=12),
            ], className="mb-4"),
            dbc.Row([
                dbc.Col([
                    html.Label('Select Title(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='title-filter',
                        options=[{'label': title, 'value': title} for title in sorted(data['Title Native Name'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Title(s)",
                    ),
                    dcc.Graph(id='rank-trend-line')
                ], width=12),
            ]),
        ]),
    ], id='tabs', active_tab='tab-overview'),

    # Footer
    html.Footer("© 2024 | Dashboard by Clarence Sai", className="text-center mt-4",
                style={'color': colors['text'], 'font-size': '14px'})
], fluid=True)

# Callbacks for interactivity
@app.callback(
    Output('total-titles', 'children'),
    Output('total-authors', 'children'),
    Output('total-publishers', 'children'),
    Output('media-type-distribution', 'figure'),
    Output('top-subjects-bar', 'figure'),
    Output('publication-year-distribution', 'figure'),
    Output('yearly-trend-chart', 'figure'),
    Output('author-heatmap', 'figure'),
    Output('rank-trend-line', 'figure'),
    Input('year-filter', 'value'),
    Input('subject-filter', 'value'),
    Input('media-filter', 'value'),
    Input('publication-date-filter', 'start_date'),
    Input('publication-date-filter', 'end_date'),
    Input('top-authors-slider', 'value'),
    Input('title-filter', 'value'),
)
def update_charts(selected_years, selected_subjects, selected_media, start_date, end_date, top_n_authors, selected_titles):
    # Filter data based on selections
    filtered_data = data.copy()
    if selected_years:
        filtered_data = filtered_data[filtered_data['Txn Calendar Year'].isin(selected_years)]
    if selected_subjects:
        filtered_data = filtered_data[filtered_data['Subject'].isin(selected_subjects)]
    if selected_media:
        filtered_data = filtered_data[filtered_data['Item Media'].isin(selected_media)]
    if start_date and end_date:
        filtered_data = filtered_data[
            (filtered_data['Title Publication Date'] >= pd.to_datetime(start_date)) &
            (filtered_data['Title Publication Date'] <= pd.to_datetime(end_date))
        ]

    # Update KPIs
    total_titles = filtered_data['Title Native Name'].nunique()
    total_authors = filtered_data['Title Author'].nunique()
    total_publishers = filtered_data['Title Publisher'].nunique()

    # Media Type Distribution Donut Chart
    if not filtered_data.empty:
        media_counts = filtered_data['Item Media'].value_counts().reset_index()
        media_counts.columns = ['Item Media', 'Count']
        fig_media = px.pie(media_counts, names='Item Media', values='Count', hole=0.5,
                           color_discrete_sequence=px.colors.sequential.Viridis,
                           title='Media Type Distribution')
        fig_media.update_traces(textposition='inside', textinfo='percent+label')
        fig_media.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_media = go.Figure()

    # Top Subjects Bar Chart
    if not filtered_data.empty:
        top_subjects = filtered_data['Subject'].value_counts().head(10).reset_index()
        top_subjects.columns = ['Subject', 'Count']
        fig_subjects = px.bar(top_subjects, x='Count', y='Subject', orientation='h',
                              title='Top 10 Subjects', color='Count',
                              color_continuous_scale=px.colors.sequential.Plasma)
        fig_subjects.update_layout(yaxis={'categoryorder': 'total ascending'},
                                   margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_subjects = go.Figure()

    # Publication Year Distribution Chart
    if not filtered_data.empty:
        publication_counts = filtered_data['Publication Year'].value_counts().reset_index()
        publication_counts.columns = ['Publication Year', 'Count']
        fig_publication_year = px.bar(publication_counts.sort_values('Publication Year'),
                                      x='Publication Year', y='Count',
                                      title='Number of Titles by Publication Year',
                                      color='Count', color_continuous_scale=px.colors.sequential.Blues)
        fig_publication_year.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_publication_year = go.Figure()

    # Yearly Trend Area Chart
    if not filtered_data.empty:
        yearly_counts = filtered_data.groupby('Txn Calendar Year').size().reset_index(name='Counts')
        fig_trend = px.area(yearly_counts, x='Txn Calendar Year', y='Counts',
                            title='Number of Titles Over Transaction Years',
                            markers=True)
        fig_trend.update_layout(xaxis_title='Transaction Year', yaxis_title='Number of Titles',
                                margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_trend = go.Figure()

    # Author Rank Trends Heatmap
    if not filtered_data.empty:
        # Get top N authors by the number of appearances
        top_authors_list = filtered_data['Title Author'].value_counts().head(top_n_authors).index.tolist()
        heatmap_data = filtered_data[filtered_data['Title Author'].isin(top_authors_list)]
        pivot_table = heatmap_data.pivot_table(values='Rank', index='Txn Calendar Year',
                                               columns='Title Author', aggfunc='mean')
        if not pivot_table.empty:
            fig_heatmap = go.Figure(data=go.Heatmap(
                z=pivot_table.values,
                x=pivot_table.columns,
                y=pivot_table.index,
                colorscale='Viridis_r',  # Reverse colorscale to reflect that lower rank is better
                hoverongaps=False))
            fig_heatmap.update_layout(title='Average Rank of Top Authors Over Years (Lower Rank is Better)',
                                      xaxis_nticks=36,
                                      margin=dict(t=50, l=25, r=25, b=25))
        else:
            fig_heatmap = go.Figure()
    else:
        fig_heatmap = go.Figure()

    # Rank Trend Line Chart
    if not filtered_data.empty:
        # If no titles are selected, default to top 5 titles by average rank
        if not selected_titles:
            top_titles = filtered_data.groupby('Title Native Name')['Rank'].mean().nsmallest(5).index.tolist()
            line_data = filtered_data[filtered_data['Title Native Name'].isin(top_titles)]
        else:
            line_data = filtered_data[filtered_data['Title Native Name'].isin(selected_titles)]
        if not line_data.empty:
            fig_rank_trend = px.line(line_data, x='Txn Calendar Year', y='Rank', color='Title Native Name',
                                     title='Rank Trend of Titles Over Years (Lower Rank is Better)', markers=True)
            fig_rank_trend.update_layout(legend=dict(orientation="h", title='Title', y=-0.2),
                                         xaxis=dict(tickmode='linear'),
                                         margin=dict(t=50, l=25, r=25, b=25))
            # Invert y-axis to reflect that lower rank is better
            fig_rank_trend.update_yaxes(autorange='reversed')
        else:
            fig_rank_trend = go.Figure()
    else:
        fig_rank_trend = go.Figure()

    return (total_titles, total_authors, total_publishers,
            fig_media, fig_subjects, fig_publication_year, fig_trend, fig_heatmap, fig_rank_trend)

# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)
