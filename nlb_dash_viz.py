import pandas as pd
from dash import Dash, dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Load the dataset
file_path = "C:\\Users\\Clarence\\Downloads\\NLB_Top100\\Top_100_OD_Titles_CY2020_to_2023.xlsx"
data = pd.read_excel(file_path, sheet_name='Sheet1')

# Data Preprocessing
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
most_common_pub_year = data['Publication Year'].mode()[0]
average_rank = round(data['Rank'].mean(), 2)

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
                html.H4(f"{most_common_pub_year}", className="card-title text-center", style={'color': colors['accent']}),
                html.P("Most Common Publication Year", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=3),
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H4(f"{average_rank}", className="card-title text-center", style={'color': colors['accent']}),
                html.P("Average Rank (Lower is Better)", className="text-center", style={'color': colors['text']})
            ])
        ], color="light", inverse=False), width=3),
    ], className="mb-5"),

    # Tabs for organizing content
    dbc.Tabs([
        # Overview Tab
        dbc.Tab(label='Overview', tab_id='tab-overview', children=[
            dbc.Row([
                dbc.Col(dcc.Graph(id='media-type-bar'), width=6),
                dbc.Col(dcc.Graph(id='top-subjects-treemap'), width=6),
            ], className="mb-4"),

            dbc.Row([
                dbc.Col(dcc.Graph(id='publication-year-histogram'), width=6),
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
                    dcc.Graph(id='author-rank-bar')
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
                    dcc.Graph(id='rank-slope-chart')
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
    Output('media-type-bar', 'figure'),
    Output('top-subjects-treemap', 'figure'),
    Output('publication-year-histogram', 'figure'),
    Output('yearly-trend-chart', 'figure'),
    Output('author-rank-bar', 'figure'),
    Output('rank-slope-chart', 'figure'),
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

    # Media Type Bar Chart
    if not filtered_data.empty:
        media_counts = filtered_data['Item Media'].value_counts().reset_index()
        media_counts.columns = ['Item Media', 'Count']
        fig_media = px.bar(media_counts, x='Item Media', y='Count',
                           color='Item Media', title='Media Type Distribution',
                           color_discrete_sequence=px.colors.qualitative.Set2)
        fig_media.update_layout(margin=dict(t=50, l=25, r=25, b=25),
                                showlegend=False)
    else:
        fig_media = go.Figure()

    # Top Subjects Treemap
    if not filtered_data.empty:
        subject_counts = filtered_data['Subject'].value_counts().reset_index()
        subject_counts.columns = ['Subject', 'Count']
        fig_subjects = px.treemap(subject_counts, path=['Subject'], values='Count',
                                  title='Top Subjects', color='Count',
                                  color_continuous_scale=px.colors.sequential.Plasma)
        fig_subjects.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_subjects = go.Figure()

    # Publication Year Histogram
    if not filtered_data.empty:
        fig_publication_year = px.histogram(filtered_data, x='Publication Year',
                                            nbins=20,
                                            title='Distribution of Publication Years',
                                            color_discrete_sequence=[colors['primary']])
        fig_publication_year.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_publication_year = go.Figure()

    # Yearly Trend Line Chart
    if not filtered_data.empty:
        yearly_counts = filtered_data.groupby('Txn Calendar Year').size().reset_index(name='Counts')
        fig_trend = px.line(yearly_counts, x='Txn Calendar Year', y='Counts',
                            title='Number of Titles Over Transaction Years', markers=True,
                            color_discrete_sequence=[colors['secondary']])
        fig_trend.update_layout(xaxis_title='Transaction Year', yaxis_title='Number of Titles',
                                margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_trend = go.Figure()

    # Author Rank Bar Chart (Lower Rank is Better)
    if not filtered_data.empty:
        top_authors_list = filtered_data['Title Author'].value_counts().index.tolist()
        # Calculate average rank for each author
        author_rank_avg = filtered_data.groupby('Title Author')['Rank'].mean().reset_index()
        # Filter top N authors based on their average rank (ascending order)
        author_rank_avg = author_rank_avg.sort_values('Rank').head(top_n_authors)
        fig_author_rank = px.bar(author_rank_avg, x='Rank', y='Title Author', orientation='h',
                                 title='Average Rank of Top Authors (Lower Rank is Better)', color='Rank',
                                 color_continuous_scale=px.colors.sequential.Viridis_r)
        fig_author_rank.update_layout(yaxis={'categoryorder': 'total ascending'},
                                      margin=dict(t=50, l=25, r=25, b=25))
    else:
        fig_author_rank = go.Figure()

    # Rank Slope Chart (Lower Rank is Better)
    if not filtered_data.empty:
        if not selected_titles:
            # Select top 5 titles based on average rank (lower is better)
            top_titles = filtered_data.groupby('Title Native Name')['Rank'].mean().nsmallest(5).index.tolist()
            slope_data = filtered_data[filtered_data['Title Native Name'].isin(top_titles)]
        else:
            slope_data = filtered_data[filtered_data['Title Native Name'].isin(selected_titles)]
        if not slope_data.empty:
            slope_data = slope_data.sort_values(['Title Native Name', 'Txn Calendar Year'])
            fig_slope = px.line(slope_data, x='Txn Calendar Year', y='Rank', color='Title Native Name',
                                markers=True, title='Rank Changes Over Years (Lower Rank is Better)')
            fig_slope.update_traces(mode='lines+markers')
            # Invert y-axis to reflect that lower rank is better
            fig_slope.update_yaxes(autorange='reversed')
            fig_slope.update_layout(legend=dict(orientation="h", title='Title', y=-0.2),
                                    xaxis=dict(tickmode='linear'),
                                    margin=dict(t=50, l=25, r=25, b=25))
        else:
            fig_slope = go.Figure()
    else:
        fig_slope = go.Figure()

    return (total_titles, total_authors, total_publishers,
            fig_media, fig_subjects, fig_publication_year, fig_trend, fig_author_rank, fig_slope)

# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)
