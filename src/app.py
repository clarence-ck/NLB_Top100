import pandas as pd
from dash import Dash, dcc, html, Input, Output, State, callback_context
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import datetime

# Load the dataset
file_path = "C:\\Users\\Clarence\\Downloads\\NLB_Top100\\Top_100_OD_Titles_CY2020_to_2023.xlsx"
data = pd.read_excel(file_path, sheet_name='Sheet1')

# Data Preprocessing
data['Title Publication Date'] = pd.to_datetime(data['Title Publication Date'], errors='coerce')
data['Publication Year'] = data['Title Publication Date'].dt.year

# Convert other columns to appropriate data types
data['Item Media'] = data['Item Media'].str.title()
data['Title Native Name'] = data['Title Native Name'].astype(str)
data['Title Author'] = data['Title Author'].astype(str)
data['Title Publisher'] = data['Title Publisher'].astype(str)
data['Txn Calendar Year'] = data['Txn Calendar Year'].astype(int)
data['Subject'] = data['Subject'].astype(str)
data['Rank'] = data['Rank'].astype(int)
data['Title Fiction Tag'] = data['Title Fiction Tag'].astype(str)

# KPI Preprocessing
total_titles = data['Title Native Name'].nunique()
total_authors = data['Title Author'].nunique()
total_publishers = data['Title Publisher'].nunique()
earliest_publication = data['Title Publication Date'].min().strftime('%Y-%m-%d')
latest_publication = data['Title Publication Date'].max().strftime('%Y-%m-%d')

# Initialize Dash app with a modern Bootstrap theme
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server  # For deployment

# Color Palette (Color-blind friendly)
colors = {
    'background': '#FFFFFF',
    'text': '#2F3E46',
    'primary': '#377eb8',
    'secondary': '#4daf4a',
    'accent': '#e41a1c',
    'info': '#984ea3',
    'warning': '#ff7f00',
    'success': '#4daf4a',
    'danger': '#a65628',
}

# App Layout
app.layout = dbc.Container([
    # Navigation Bar
    dbc.Navbar(
        [
            html.A(
                dbc.Row(
                    [
                        dbc.Col(
                            html.Img(
                                src="https://fundit.fr/sites/default/files/styles/max_650x650/public/institutions/capture-decran-2023-06-30-143348.png?itok=dP4xFwqc",
                                height="40px"
                            )
                        ),
                        dbc.Col(
                            dbc.NavbarBrand(
                                "Top 100 OverDrive Titles Dashboard (2020 - 2023)", className="ml-2"
                            )
                        ),
                    ],
                    align="center",
                    className="g-0",
                ),
                href="#",
                style={"textDecoration": "none"},
            ),
        ],
        color=colors['primary'],
        dark=True,
        className="mb-4",
    ),

    # Header and Description
    dbc.Row([
        dbc.Col([
            html.H2("Explore the Trends and Insights of the Top 100 NLB OD Titles"),
            html.P(
                "Dive into the interactive dashboard to uncover patterns and trends in the most popular OverDrive titles from 2020 to 2023. Use the filters below to customize your view."
            ),
        ], width=12)
    ], className="mb-4"),

    # Filters
    dbc.Card(
        dbc.CardBody([
            dbc.Row([
                dbc.Col([
                    html.Label('Transaction Year(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='year-filter',
                        options=[{'label': str(year), 'value': year} for year in sorted(data['Txn Calendar Year'].unique())],
                        value=sorted(data['Txn Calendar Year'].unique()),
                        multi=True,
                        placeholder="Select Year(s)",
                    )
                ], md=2),

                dbc.Col([
                    html.Label('Subject(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='subject-filter',
                        options=[{'label': subj, 'value': subj} for subj in sorted(data['Subject'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Subject(s)",
                        style={'width': '100%'}
                    )
                ], md=4),

                dbc.Col([
                    html.Label('Author(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='author-filter',
                        options=[{'label': author, 'value': author} for author in sorted(data['Title Author'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Author(s)",
                        style={'width': '100%'}
                    )
                ], md=3),

                dbc.Col([
                    html.Label('Publisher(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='publisher-filter',
                        options=[{'label': publisher, 'value': publisher} for publisher in sorted(data['Title Publisher'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Publisher(s)",
                        style={'width': '100%'}
                    )
                ], md=3),
            ], className="mb-3"),

            dbc.Row([
                dbc.Col([
                    html.Label('Category:', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='fiction-filter',
                        options=[
                            {'label': 'Fiction', 'value': 'Yes'},
                            {'label': 'Non-Fiction', 'value': 'No'}
                        ],
                        value=['Yes', 'No'],  # Default to both options selected
                        multi=True,
                        placeholder="Select Category",
                        style={'width': '100%'}
                    )
                ], md=3),

                dbc.Col([
                    html.Label('Media Type(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='media-filter',
                        options=[{'label': media, 'value': media} for media in sorted(data['Item Media'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Media Type(s)",
                        style={'width': '100%'}
                    )
                ], md=3),

                dbc.Col([
                    html.Label('Publication Date Range:', style={'font-weight': 'bold'}),
                    dcc.RangeSlider(
                        id='publication-date-slider',
                        min=data['Publication Year'].min(),
                        max=data['Publication Year'].max(),
                        value=[data['Publication Year'].min(), data['Publication Year'].max()],
                        marks={year: str(year) for year in range(data['Publication Year'].min(), data['Publication Year'].max()+1, 2)},
                        tooltip={"placement": "bottom", "always_visible": False},
                    )
                ], md=6),
            ]),
        ]),
        className="mb-4",
    ),

    # KPI Cards
    dbc.Row([
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Unique Titles", className="card-title"),
                html.H3(id='total-titles', className="card-text"),
            ])
        ], color=colors['info'], inverse=True), md=3),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Unique Authors", className="card-title"),
                html.H3(id='total-authors', className="card-text"),
            ])
        ], color=colors['success'], inverse=True), md=3),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Unique Publishers", className="card-title"),
                html.H3(id='total-publishers', className="card-text"),
            ])
        ], color=colors['warning'], inverse=True), md=3),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.H5("Time Span", className="card-title"),
                html.H3(id='time-span', className="card-text"),
            ])
        ], color=colors['danger'], inverse=True), md=3),
    ], className="mb-5"),

    # Tabs for organizing content
    dbc.Tabs([
        # Overview Tab
        dbc.Tab(label='Overview', tab_id='tab-overview', children=[
            # First Row
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-media-type-donut',
                        type='default',
                        children=dcc.Graph(id='media-type-donut')
                    ),
                    md=6
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-category-distribution-donut',
                        type='default',
                        children=dcc.Graph(id='category-distribution-donut')
                    ),
                    md=6
                ),
            ], className="mb-4"),

            # Second Row
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-top-publishers-bar',
                        type='default',
                        children=dcc.Graph(id='top-publishers-bar')
                    ),
                    md=6
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-top-authors-bar',
                        type='default',
                        children=dcc.Graph(id='top-authors-bar')
                    ),
                    md=6
                ),
            ], className="mb-4"),

            # Third Row
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-publication-year-stacked-bar',
                        type='default',
                        children=dcc.Graph(id='publication-year-stacked-bar')
                    ),
                    md=6
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-custom-chart',
                        type='default',
                        children=dcc.Graph(id='custom-chart')
                    ),
                    md=6
                ),
            ], className="mb-4"),
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
                    dcc.Loading(
                        id='loading-author-heatmap',
                        type='default',
                        children=dcc.Graph(id='author-heatmap')
                    ),
                ], md=12),
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
                    dcc.Loading(
                        id='loading-rank-trend-line',
                        type='default',
                        children=dcc.Graph(id='rank-trend-line')
                    ),
                ], md=12),
            ]),
        ]),
    ], id='tabs', active_tab='tab-overview'),

    # Footer
    html.Footer(
        [
            html.Hr(),
            html.P("© 2024 | Dashboard by Clarence Sai", className="text-center"),
        ]
    )
], fluid=True, style={'backgroundColor': colors['background']})

# Helper functions to create figures
def create_media_type_donut_chart(filtered_data):
    if not filtered_data.empty:
        media_counts = filtered_data['Item Media'].value_counts().reset_index()
        media_counts.columns = ['Item Media', 'Count']
        fig = px.pie(media_counts, names='Item Media', values='Count', hole=0.4,
                     title='Media Type Distribution',
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        return go.Figure()

def create_category_distribution_donut_chart(filtered_data):
    if not filtered_data.empty:
        category_counts = filtered_data['Title Fiction Tag'].value_counts().reset_index()
        category_counts.columns = ['Category', 'Count']
        category_counts['Category'] = category_counts['Category'].map({'Yes': 'Fiction', 'No': 'Non-Fiction'})
        fig = px.pie(category_counts, names='Category', values='Count', hole=0.4,
                     title='Category Distribution',
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        return go.Figure()

def create_top_publishers_bar_chart(filtered_data):
    if not filtered_data.empty:
        top_publishers = filtered_data['Title Publisher'].value_counts().head(10).reset_index()
        top_publishers.columns = ['Publisher', 'Count']
        fig = px.bar(top_publishers, x='Count', y='Publisher', orientation='h',
                     title='Top 10 Publishers',
                     color='Count', color_continuous_scale=px.colors.sequential.Viridis)
        fig.update_layout(yaxis={'categoryorder': 'total ascending'},
                          margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        return go.Figure()

def create_top_authors_bar_chart(filtered_data):
    if not filtered_data.empty:
        top_authors = filtered_data['Title Author'].value_counts().head(10).reset_index()
        top_authors.columns = ['Author', 'Count']
        fig = px.bar(top_authors, x='Count', y='Author', orientation='h',
                     title='Top 10 Authors',
                     color='Count', color_continuous_scale=px.colors.sequential.Plasma)
        fig.update_layout(yaxis={'categoryorder': 'total ascending'},
                          margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        return go.Figure()

def create_publication_year_stacked_bar_chart(filtered_data):
    if not filtered_data.empty:
        df = filtered_data.copy()
        df['Category'] = df['Title Fiction Tag'].map({'Yes': 'Fiction', 'No': 'Non-Fiction'})
        publication_counts = df.groupby(['Publication Year', 'Category']).size().reset_index(name='Count')
        fig = px.bar(publication_counts, x='Publication Year', y='Count', color='Category',
                     title='Number of Titles by Publication Year',
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        return go.Figure()

def create_transaction_year_media_type_chart(filtered_data):
    if not filtered_data.empty:
        counts = filtered_data.groupby(['Txn Calendar Year', 'Item Media']).size().reset_index(name='Count')
        fig = px.bar(counts, x='Txn Calendar Year', y='Count', color='Item Media',
                     title='Number of Titles by Transaction Year and Media Type',
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        return go.Figure()

def create_author_heatmap(filtered_data, top_n_authors):
    if not filtered_data.empty:
        top_authors_list = filtered_data['Title Author'].value_counts().head(top_n_authors).index.tolist()
        heatmap_data = filtered_data[filtered_data['Title Author'].isin(top_authors_list)]
        pivot_table = heatmap_data.pivot_table(values='Rank', index='Txn Calendar Year',
                                               columns='Title Author', aggfunc='mean')
        if not pivot_table.empty:
            fig = px.imshow(pivot_table.T, aspect='auto',
                            color_continuous_scale='Viridis_r',
                            title='Average Rank of Top Authors Over Years (Lower Rank is Better)')
            fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
            return fig
        else:
            return go.Figure()
    else:
        return go.Figure()

def create_rank_trend_line_chart(filtered_data, selected_titles):
    if not filtered_data.empty:
        if selected_titles:
            line_data = filtered_data[filtered_data['Title Native Name'].isin(selected_titles)]
            if not line_data.empty:
                fig = px.line(line_data, x='Txn Calendar Year', y='Rank', color='Title Native Name',
                              title='Rank Trend of Titles Over Years (Lower Rank is Better)', markers=True)
                fig.update_layout(legend=dict(orientation="h", title='Title', y=-0.2),
                                  xaxis=dict(tickmode='linear'),
                                  margin=dict(t=50, l=25, r=25, b=25))
                fig.update_yaxes(autorange='reversed')
                return fig
            else:
                fig = go.Figure()
                fig.add_annotation(
                    x=0.5, y=0.5,
                    text="Selected titles not found in the filtered data.",
                    showarrow=False,
                    xref="paper", yref="paper",
                    font=dict(size=16),
                    xanchor='center', yanchor='middle'
                )
                fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
                return fig
        else:
            # No titles selected, return an empty figure with a message
            fig = go.Figure()
            fig.add_annotation(
                x=0.5, y=0.5,
                text="Please select at least one title to display the rank trend.",
                showarrow=False,
                xref="paper", yref="paper",
                font=dict(size=16),
                xanchor='center', yanchor='middle'
            )
            fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
            return fig
    else:
        return go.Figure()

# Callbacks for interactivity
@app.callback(
    Output('total-titles', 'children'),
    Output('total-authors', 'children'),
    Output('total-publishers', 'children'),
    Output('time-span', 'children'),
    Output('media-type-donut', 'figure'),
    Output('category-distribution-donut', 'figure'),
    Output('top-publishers-bar', 'figure'),
    Output('top-authors-bar', 'figure'),
    Output('publication-year-stacked-bar', 'figure'),
    Output('custom-chart', 'figure'),
    Output('author-heatmap', 'figure'),
    Output('rank-trend-line', 'figure'),
    Input('year-filter', 'value'),
    Input('subject-filter', 'value'),
    Input('media-filter', 'value'),
    Input('publication-date-slider', 'value'),
    Input('top-authors-slider', 'value'),
    Input('title-filter', 'value'),
    Input('author-filter', 'value'),
    Input('publisher-filter', 'value'),
    Input('fiction-filter', 'value'),
)
def update_charts(selected_years, selected_subjects, selected_media,
                  publication_year_range, top_n_authors, selected_titles,
                  selected_authors, selected_publishers, selected_fiction):
    # Filter data based on selections
    filtered_data = data

    if selected_years:
        filtered_data = filtered_data.loc[filtered_data['Txn Calendar Year'].isin(selected_years)]

    if selected_subjects:
        filtered_data = filtered_data.loc[filtered_data['Subject'].isin(selected_subjects)]

    if selected_media:
        filtered_data = filtered_data.loc[filtered_data['Item Media'].isin(selected_media)]

    if publication_year_range:
        start_year, end_year = publication_year_range
        filtered_data = filtered_data.loc[
            (filtered_data['Publication Year'] >= start_year) &
            (filtered_data['Publication Year'] <= end_year)
        ]

    if selected_authors:
        filtered_data = filtered_data.loc[filtered_data['Title Author'].isin(selected_authors)]

    if selected_publishers:
        filtered_data = filtered_data.loc[filtered_data['Title Publisher'].isin(selected_publishers)]

    if selected_fiction:
        filtered_data = filtered_data.loc[filtered_data['Title Fiction Tag'].isin(selected_fiction)]

    # Update KPIs
    total_titles = filtered_data['Title Native Name'].nunique()
    total_authors = filtered_data['Title Author'].nunique()
    total_publishers = filtered_data['Title Publisher'].nunique()

    # Time Span
    min_year = filtered_data['Publication Year'].min()
    max_year = filtered_data['Publication Year'].max()
    time_span = f"{int(min_year)} - {int(max_year)}" if pd.notnull(min_year) and pd.notnull(max_year) else "N/A"

    # Create charts using helper functions
    fig_media_donut = create_media_type_donut_chart(filtered_data)
    fig_category_donut = create_category_distribution_donut_chart(filtered_data)
    fig_top_publishers_bar = create_top_publishers_bar_chart(filtered_data)
    fig_top_authors_bar = create_top_authors_bar_chart(filtered_data)
    fig_publication_year_stacked_bar = create_publication_year_stacked_bar_chart(filtered_data)
    fig_custom_chart = create_transaction_year_media_type_chart(filtered_data)  # Custom chart

    fig_heatmap = create_author_heatmap(filtered_data, top_n_authors)
    fig_rank_trend = create_rank_trend_line_chart(filtered_data, selected_titles)

    return (total_titles, total_authors, total_publishers, time_span,
            fig_media_donut, fig_category_donut, fig_top_publishers_bar, fig_top_authors_bar,
            fig_publication_year_stacked_bar, fig_custom_chart,
            fig_heatmap, fig_rank_trend)

# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)
