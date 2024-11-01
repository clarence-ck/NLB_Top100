import pandas as pd
from dash import Dash, dcc, html, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go

# Load the dataset
file_path = "C:\\Users\\Clarence\\Downloads\\NLB_Top100\\Top_100_OD_Titles_CY2020_to_2023.xlsx"
data = pd.read_excel(file_path, sheet_name='Sheet1')

# Data Preprocessing
data['Title Publication Date'] = pd.to_datetime(data['Title Publication Date'], errors='coerce')
data['Publication Year'] = data['Title Publication Date'].dt.year
data['Item Media'] = data['Item Media'].str.title()
data['Title Native Name'] = data['Title Native Name'].astype(str)
data['Title Author'] = data['Title Author'].astype(str)
data['Title Publisher'] = data['Title Publisher'].astype(str)
data['Txn Calendar Year'] = data['Txn Calendar Year'].astype(int)
data['Subject'] = data['Subject'].astype(str)
data['Rank'] = data['Rank'].astype(int)
data['Title Fiction Tag'] = data['Title Fiction Tag'].astype(str)

# Initialize Dash app with Bootstrap theme
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server  # For deployment

# Define colors
colors = {
    'background': '#F8F9FA',
    'text': '#343A40',
    'primary': '#1E90FF',        # DodgerBlue
    'secondary': '#20C997',      # SeaGreen
    'accent': '#FF6B6B',         # Light Red
    'info': '#6F42C1',           # Purple
    'warning': '#FFC107',        # Amber
    'success': '#28A745',        # Green
    'danger': '#DC3545',         # Red
    'light': '#E9ECEF',
    'dark': '#343A40',
}

# Include Bootstrap Icons
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <!-- Bootstrap Icons CDN -->
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css">
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

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
            html.H2("Explore the Trends and Insights of the Top 100 NLB OD Titles", className="text-center"),
            html.P(
                "Dive into the interactive dashboard to uncover patterns and trends in the most popular OverDrive titles from 2020 to 2023. Use the filters below to customize your view.",
                className="text-center"
            ),
        ], width=12)
    ], className="mb-4"),

    # Filters in a Card
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
                ], md=2, sm=6, xs=12),
                dbc.Col([
                    html.Label('Subject(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='subject-filter',
                        options=[{'label': subj, 'value': subj} for subj in sorted(data['Subject'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Subject(s)"
                    )
                ], md=4, sm=6, xs=12),
                dbc.Col([
                    html.Label('Author(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='author-filter',
                        options=[{'label': author, 'value': author} for author in sorted(data['Title Author'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Author(s)"
                    )
                ], md=3, sm=6, xs=12),
                dbc.Col([
                    html.Label('Publisher(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='publisher-filter',
                        options=[{'label': publisher, 'value': publisher} for publisher in sorted(data['Title Publisher'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Publisher(s)"
                    )
                ], md=3, sm=6, xs=12),
            ], className="mb-3"),
            dbc.Row([
                dbc.Col([
                    html.Label('Category:', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='fiction-filter',
                        options=[{'label': 'Fiction', 'value': 'Yes'}, {'label': 'Non-Fiction', 'value': 'No'}],
                        value=['Yes', 'No'],
                        multi=True,
                        placeholder="Select Category"
                    )
                ], md=3, sm=6, xs=12),
                dbc.Col([
                    html.Label('Media Type(s):', style={'font-weight': 'bold'}),
                    dcc.Dropdown(
                        id='media-filter',
                        options=[{'label': media, 'value': media} for media in sorted(data['Item Media'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Media Type(s)"
                    )
                ], md=3, sm=6, xs=12),
                dbc.Col([
                    html.Label('Select Publication Start Date:', style={'font-weight': 'bold', 'display': 'block'}),
                    dcc.DatePickerSingle(
                        id='publication-start-date-filter',
                        min_date_allowed=data['Title Publication Date'].min(),
                        max_date_allowed=data['Title Publication Date'].max(),
                        date=data['Title Publication Date'].min(),
                        display_format='YYYY-MM-DD',
                    ),
                ], md=3, sm=6, xs=12),
                dbc.Col([
                    html.Label('Select Publication End Date:', style={'font-weight': 'bold', 'display': 'block'}),
                    dcc.DatePickerSingle(
                        id='publication-end-date-filter',
                        min_date_allowed=data['Title Publication Date'].min(),
                        max_date_allowed=data['Title Publication Date'].max(),
                        date=data['Title Publication Date'].max(),
                        display_format='YYYY-MM-DD',
                    ),
                ], md=3, sm=6, xs=12),
            ]),
        ]),
        className="mb-4",
        style={'box-shadow': '0 4px 8px rgba(0,0,0,0.1)', 'padding': '20px'}
    ),

    # KPIs
    dbc.Row([
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-book", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Unique Titles", className="card-title mt-2"),
                html.H3(id='total-titles', className="card-text"),
            ])
        ], color=colors['info'], inverse=True, className="text-center shadow-sm", style={'height': '150px'}), md=2, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-person", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Unique Authors", className="card-title mt-2"),
                html.H3(id='total-authors', className="card-text"),
            ])
        ], color=colors['success'], inverse=True, className="text-center shadow-sm", style={'height': '150px'}), md=2, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-people", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Unique Publishers", className="card-title mt-2"),
                html.H3(id='total-publishers', className="card-text"),
            ])
        ], color=colors['warning'], inverse=True, className="text-center shadow-sm", style={'height': '150px'}), md=2, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-calendar", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Earliest Publication Date", className="card-title mt-2"),
                html.H3(id='earliest-publication', className="card-text"),
            ])
        ], color=colors['primary'], inverse=True, className="text-center shadow-sm", style={'height': '150px'}), md=3, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-calendar-check", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Latest Publication Date", className="card-title mt-2"),
                html.H3(id='latest-publication', className="card-text"),
            ])
        ], color=colors['secondary'], inverse=True, className="text-center shadow-sm", style={'height': '150px'}), md=3, sm=6, xs=12),
    ], className="mb-4"),

    # Tabs for organizing content
    dbc.Tabs([
        # Overview Tab
        dbc.Tab(label='Overview', tab_id='tab-overview', children=[
            # First Row: Media Type and Category Donut Charts
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

            # Second Row: OverDrive Distribution Treemap
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-overdrive-distribution',  # New ID
                        type='default',
                        children=dcc.Graph(id='overdrive-distribution')  # New ID
                    ),
                    md=12  # Full width
                ),
            ], className="mb-4"),

            # Third Row: Top 10 Publishers and Top 10 Authors Bar Charts
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-top-publishers',
                        type='default',
                        children=dcc.Graph(id='top-publishers-bar')
                    ),
                    md=6
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-top-authors',
                        type='default',
                        children=dcc.Graph(id='top-authors-bar')
                    ),
                    md=6
                ),
            ], className="mb-4"),

            # Fourth Row: Publication Year Stacked Bar and Custom Chart
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

        # Rank Trend Analysis Tab
        dbc.Tab(label='Rank Trend Analysis', tab_id='tab-detailed', children=[
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

            # Explanation for Average Rank of Top Authors
            dbc.Row([
                dbc.Col([
                    html.P(
                        "The heatmap above displays the average rank of the top authors over the years. "
                        "The average rank is calculated by taking the mean rank of each author's titles for each transaction year. "
                        "A lower rank indicates better performance."
                    )
                ], md=12),
            ], className="mb-2"),

            # Rank Trend Line Chart
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
        fig = go.Figure()
        fig.update_layout(
            title="Media Type Distribution",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

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
        fig = go.Figure()
        fig.update_layout(
            title="Category Distribution",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

def create_overdrive_distribution_treemap(filtered_data):
    if not filtered_data.empty:
        # Prepare the hierarchical data
        treemap_data = filtered_data.groupby(['Title Publisher', 'Title Author', 'Subject']).size().reset_index(name='Count')

        fig = px.treemap(
            treemap_data,
            path=['Title Publisher', 'Title Author', 'Subject'],
            values='Count',
            title='OverDrive Distribution',
            color='Count',
            color_continuous_scale='Viridis',
            hover_data={'Count': True},
            labels={'Title Publisher': 'Publisher', 'Title Author': 'Author', 'Subject': 'Subject'}
        )
        fig.update_traces(
            hovertemplate='<b>%{label}</b><br>Count: %{value}<extra></extra>'
        )
        fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
        return fig
    else:
        fig = go.Figure()
        fig.update_layout(
            title="OverDrive Distribution",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

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
        fig = go.Figure()
        fig.update_layout(
            title="Number of Titles by Publication Year",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

def create_transaction_year_media_type_chart(filtered_data):
    if not filtered_data.empty:
        # Group data by Transaction Year and Item Media, then count the occurrences
        counts = filtered_data.groupby(['Txn Calendar Year', 'Item Media']).size().reset_index(name='Count')
        
        # Create the bar chart
        fig = px.bar(
            counts,
            x='Txn Calendar Year',
            y='Count',
            color='Item Media',
            title='Number of Titles by Transaction Year and Media Type',
            color_discrete_sequence=px.colors.qualitative.Pastel,
            labels={
                'Txn Calendar Year': 'Transaction Year',
                'Count': 'Number of Titles',
                'Item Media': 'Media Type'
            }
        )
        
        # Update the x-axis to show only the years 2020 to 2023 without sub-units
        fig.update_xaxes(
            tickmode='linear',
            tick0=2020,
            dtick=1,
            range=[2019.5, 2023.5],  # Slightly beyond to ensure full bars are visible
            title='Transaction Year'
        )
        
        # Update the y-axis title
        fig.update_yaxes(title='Number of Titles')
        
        # Enhance overall layout aesthetics
        fig.update_layout(
            margin=dict(t=50, l=50, r=25, b=50),
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor='LightGray')
        )
        return fig
    else:
        # Create an empty figure with a "No data available" message
        fig = go.Figure()
        fig.update_layout(
            title="Number of Titles by Transaction Year and Media Type",
            annotations=[
                dict(
                    text="No data available",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(size=20)
                )
            ],
            xaxis=dict(
                visible=False
            ),
            yaxis=dict(
                visible=False
            )
        )
        return fig

def create_top_publishers_bar_chart(filtered_data):
    if not filtered_data.empty:
        top_publishers = filtered_data['Title Publisher'].value_counts().head(10).reset_index()
        top_publishers.columns = ['Title Publisher', 'Count']
        fig = px.bar(
            top_publishers,
            x='Count',
            y='Title Publisher',
            orientation='h',
            title='Top 10 Publishers by Number of Titles',
            labels={'Count': 'Number of Titles', 'Title Publisher': 'Publisher'},
            color='Count',
            color_continuous_scale='Blues',
            height=500
        )
        fig.update_layout(
            margin=dict(t=50, l=100, r=25, b=50),
            yaxis=dict(autorange="reversed"),  # Highest count on top
            coloraxis_showscale=True
        )
        return fig
    else:
        fig = go.Figure()
        fig.update_layout(
            title="Top 10 Publishers by Number of Titles",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

def create_top_authors_bar_chart(filtered_data):
    if not filtered_data.empty:
        top_authors = filtered_data['Title Author'].value_counts().head(10).reset_index()
        top_authors.columns = ['Title Author', 'Count']
        fig = px.bar(
            top_authors,
            x='Count',
            y='Title Author',
            orientation='h',
            title='Top 10 Authors by Number of Titles',
            labels={'Count': 'Number of Titles', 'Title Author': 'Author'},
            color='Count',
            color_continuous_scale='Greens',
            height=500
        )
        fig.update_layout(
            margin=dict(t=50, l=100, r=25, b=50),
            yaxis=dict(autorange="reversed"),  # Highest count on top
            coloraxis_showscale=True
        )
        return fig
    else:
        fig = go.Figure()
        fig.update_layout(
            title="Top 10 Authors by Number of Titles",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

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
            fig.update_layout(
                xaxis_title='Transaction Calendar Year',  # Updated X-axis label
                yaxis_title='Author',                      # Added Y-axis label for clarity
                xaxis=dict(
                    tickmode='linear',
                    tick0=2020,
                    dtick=1,
                    range=[2019.5, 2023.5],  
                ),
                margin=dict(t=100, l=150, r=25, b=50)     # Adjusted margins for better layout
            )
            return fig
        else:
            fig = go.Figure()
            fig.update_layout(
                title="Average Rank of Top Authors Over Years",
                annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
            )
            return fig
    else:
        fig = go.Figure()
        fig.update_layout(
            title="Average Rank of Top Authors Over Years",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)]
        )
        return fig

# Updated Rank Trend Line Chart Function
def create_rank_trend_line_chart(filtered_data, selected_titles):
    if not filtered_data.empty:
        if selected_titles:
            line_data = filtered_data[filtered_data['Title Native Name'].isin(selected_titles)]
            if not line_data.empty:
                fig = px.line(
                    line_data,
                    x='Txn Calendar Year',
                    y='Rank',
                    color='Title Native Name',
                    title='Rank Trend of Titles Over Years (Lower Rank is Better)',
                    markers=True
                )
                
                # Update marker properties to enhance visibility
                fig.update_traces(
                    marker=dict(size=10),
                    cliponaxis=False  # Prevent markers from being clipped
                )
                
                # Update layout with extended x-axis range and increased margins
                fig.update_layout(
                    legend=dict(
                        orientation="h",
                        title='Title',
                        y=-0.2
                    ),
                    xaxis=dict(
                        tickmode='linear',
                        tick0=2020,
                        dtick=1,
                        range=[2020, 2023],  # Added padding to x-axis
                        title='Transaction Calendar Year'
                    ),
                    yaxis=dict(
                        title='Rank',
                        autorange='reversed'  # Ensures lower ranks are higher on the axis
                    ),
                    margin=dict(
                        t=100,
                        l=60,   # Increased left margin
                        r=60,   # Increased right margin
                        b=100
                    )
                )
                
                return fig
            else:
                fig = go.Figure()
                fig.update_layout(
                    title="Rank Trend of Titles Over Years",
                    annotations=[
                        dict(
                            text="Selected titles not found in the filtered data.",
                            x=0.5,
                            y=0.5,
                            showarrow=False
                        )
                    ]
                )
                return fig
        else:
            # No titles selected, return an empty figure with a message
            fig = go.Figure()
            fig.add_annotation(
                x=0.5,
                y=0.5,
                text="Please select at least one title to display the rank trend.",
                showarrow=False,
                xref="paper",
                yref="paper",
                font=dict(size=16),
                xanchor='center',
                yanchor='middle'
            )
            fig.update_layout(
                title="Rank Trend of Titles Over Years",
                margin=dict(t=100, l=50, r=25, b=100)
            )
            return fig
    else:
        # No data available after filtering
        fig = go.Figure()
        fig.update_layout(
            title="Rank Trend of Titles Over Years",
            annotations=[
                dict(
                    text="No data available",
                    x=0.5,
                    y=0.5,
                    showarrow=False
                )
            ]
        )
        return fig

# Callbacks for interactivity
@app.callback(
    [
        Output('total-titles', 'children'),
        Output('total-authors', 'children'),
        Output('total-publishers', 'children'),
        Output('earliest-publication', 'children'),   # New Output
        Output('latest-publication', 'children'),     # New Output
        Output('media-type-donut', 'figure'),
        Output('category-distribution-donut', 'figure'),
        Output('overdrive-distribution', 'figure'),   # New Output
        Output('top-publishers-bar', 'figure'),       # New Output
        Output('top-authors-bar', 'figure'),          # New Output
        Output('publication-year-stacked-bar', 'figure'),
        Output('custom-chart', 'figure'),
        Output('author-heatmap', 'figure'),
        Output('rank-trend-line', 'figure'),
    ],
    [
        Input('year-filter', 'value'),
        Input('subject-filter', 'value'),
        Input('media-filter', 'value'),
        Input('publication-start-date-filter', 'date'),  # New Input
        Input('publication-end-date-filter', 'date'),    # New Input
        Input('top-authors-slider', 'value'),
        Input('title-filter', 'value'),
        Input('author-filter', 'value'),
        Input('publisher-filter', 'value'),
        Input('fiction-filter', 'value'),
    ]
)
def update_charts(selected_years, selected_subjects, selected_media,
                  publication_start_date, publication_end_date, top_n_authors,
                  selected_titles, selected_authors, selected_publishers, selected_fiction):
    # Filter data based on selections
    filtered_data = data.copy()

    if selected_years:
        filtered_data = filtered_data[filtered_data['Txn Calendar Year'].isin(selected_years)]

    if selected_subjects:
        filtered_data = filtered_data[filtered_data['Subject'].isin(selected_subjects)]

    if selected_media:
        filtered_data = filtered_data[filtered_data['Item Media'].isin(selected_media)]

    if publication_start_date:
        filtered_data = filtered_data[filtered_data['Title Publication Date'] >= pd.to_datetime(publication_start_date)]

    if publication_end_date:
        filtered_data = filtered_data[filtered_data['Title Publication Date'] <= pd.to_datetime(publication_end_date)]

    if selected_authors:
        filtered_data = filtered_data[filtered_data['Title Author'].isin(selected_authors)]

    if selected_publishers:
        filtered_data = filtered_data[filtered_data['Title Publisher'].isin(selected_publishers)]

    if selected_fiction:
        filtered_data = filtered_data[filtered_data['Title Fiction Tag'].isin(selected_fiction)]

    # Update KPIs
    total_titles = filtered_data['Title Native Name'].nunique()
    total_authors = filtered_data['Title Author'].nunique()
    total_publishers = filtered_data['Title Publisher'].nunique()

    # Calculate Earliest and Latest Publication Dates
    earliest_date = filtered_data['Title Publication Date'].min()
    latest_date = filtered_data['Title Publication Date'].max()

    # Format dates as strings for display, handle missing values
    earliest_publication = earliest_date.strftime('%Y-%m-%d') if pd.notnull(earliest_date) else "N/A"
    latest_publication = latest_date.strftime('%Y-%m-%d') if pd.notnull(latest_date) else "N/A"

    # Create charts using helper functions
    fig_media_donut = create_media_type_donut_chart(filtered_data)
    fig_category_donut = create_category_distribution_donut_chart(filtered_data)
    fig_overdrive_distribution = create_overdrive_distribution_treemap(filtered_data)  # New
    fig_top_publishers = create_top_publishers_bar_chart(filtered_data)               # New
    fig_top_authors = create_top_authors_bar_chart(filtered_data)                     # New
    fig_publication_year_stacked_bar = create_publication_year_stacked_bar_chart(filtered_data)
    fig_custom_chart = create_transaction_year_media_type_chart(filtered_data)  # Custom chart

    fig_heatmap = create_author_heatmap(filtered_data, top_n_authors)
    fig_rank_trend = create_rank_trend_line_chart(filtered_data, selected_titles)

    return (
        total_titles,           
        total_authors,         
        total_publishers,       
        earliest_publication,   
        latest_publication,     
        fig_media_donut,
        fig_category_donut,
        fig_overdrive_distribution,  # New
        fig_top_publishers,          # New
        fig_top_authors,             # New
        fig_publication_year_stacked_bar,
        fig_custom_chart,
        fig_heatmap,
        fig_rank_trend
    )

# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)