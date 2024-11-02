import pandas as pd
from dash import Dash, dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go

# Load the dataset
file_path = "https://github.com/clarence-ck/NLB_Top100/raw/refs/heads/main/Top_100_OD_Titles_CY2020_to_2023.xlsx"
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

# Define the years for which heatmaps will be created
HEATMAP_YEARS = sorted(data['Txn Calendar Year'].unique())

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

# Helper function to generate tooltip string with truncated titles
def generate_ranks_str(titles, ranks, max_display=10, max_title_length=55):
    """
    Generates a tooltip string in the format "Title: Rank" separated by line breaks.
    
    Parameters:
    - titles (list): List of book titles.
    - ranks (list): List of corresponding ranks.
    - max_display (int): Maximum number of title-rank pairs to display in the tooltip.
    - max_title_length (int): Maximum number of characters for each title.
    
    Returns:
    - str: Formatted tooltip string.
    """
    # Validate that both titles and ranks are lists
    if not isinstance(titles, list) or not isinstance(ranks, list):
        return "No rank data available."
    
    # Determine the number of pairs to process
    num_pairs = min(len(titles), len(ranks), max_display)
    
    # Initialize an empty list to store "Title: Rank" strings
    rank_entries = []
    
    for i in range(num_pairs):
        title = titles[i] if pd.notnull(titles[i]) else "N/A"
        rank = ranks[i] if pd.notnull(ranks[i]) else "N/A"
        
        # Truncate the title if it's too long
        if len(title) > max_title_length:
            title = title[:max_title_length-3] + "..."
        
        rank_entries.append(f"{title}: {rank}")
    
    # Indicate if there are additional titles or ranks not displayed
    if len(titles) > max_display or len(ranks) > max_display:
        rank_entries.append("...")
    
    # Join the entries with HTML line breaks
    return '<br>'.join(rank_entries)

# App Layout
app.layout = dbc.Container([
    # Navigation Bar
    dbc.Navbar(
        dbc.Container([
            html.A(
                dbc.Row(
                    [
                        dbc.Col(
                            html.Img(
                                src="https://fundit.fr/sites/default/files/styles/max_650x650/public/institutions/capture-decran-2023-06-30-143348.png?itok=dP4xFwqc",
                                height="70px",  # Increased height from 40px to 70px
                                className="d-inline-block align-bottom"
                            )
                        ),
                        dbc.Col(
                            dbc.NavbarBrand(
                                "Top 100 OverDrive Titles Dashboard (2020 - 2023)", 
                                className="ms-2"
                            )
                        ),
                    ],
                    align="center",
                    className="g-0",
                ),
                href="#",
                style={"textDecoration": "none"},
            ),
        ]),
        color=colors['primary'],
        dark=True,
        fixed="top",
        className="mb-4",
    ),

    # Spacer for fixed Navbar
    html.Div(style={'height': '140px'}),  # Increased height to accommodate larger Navbar image

    # Header and Description
    dbc.Row([
        dbc.Col([
            html.H2("Explore the Trends and Insights of the Top 100 NLB OD Titles", className="text-center mb-3"),
            html.P(
                "Dive into the interactive dashboard to uncover patterns and trends in the most popular OverDrive titles from 2020 to 2023. Use the filters below to customize your view.",
                className="text-center"
            ),
        ], width=12)
    ], className="mb-4"),

    # Filters in a Card
    dbc.Card(
        dbc.CardBody([
            # First Row of Filters
            dbc.Row([
                dbc.Col([
                    html.Label('Transaction Year(s):', className="fw-bold"),
                    dcc.Dropdown(
                        id='year-filter',
                        options=[{'label': str(year), 'value': year} for year in sorted(data['Txn Calendar Year'].unique())],
                        value=sorted(data['Txn Calendar Year'].unique()),
                        multi=True,
                        placeholder="Select Year(s)",
                    )
                ], md=3, sm=6, xs=12),
                dbc.Col([
                    html.Label('Subject(s):', className="fw-bold"),
                    dcc.Dropdown(
                        id='subject-filter',
                        options=[{'label': subj, 'value': subj} for subj in sorted(data['Subject'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Subject(s)",
                        style={'height': 'auto', 'minHeight': '40px'},
                        clearable=True,
                        optionHeight=50,
                    )
                ], md=4, sm=6, xs=12),

                dbc.Col([
                    html.Label('Author(s):', className="fw-bold"),
                    dcc.Dropdown(
                        id='author-filter',
                        options=[{'label': author, 'value': author} for author in sorted(data['Title Author'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Author(s)"
                    )
                ], md=2, sm=6, xs=12),
                dbc.Col([
                    html.Label('Publisher(s):', className="fw-bold"),
                    dcc.Dropdown(
                        id='publisher-filter',
                        options=[{'label': publisher, 'value': publisher} for publisher in sorted(data['Title Publisher'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Publisher(s)"
                    )
                ], md=2, sm=6, xs=12),
            ], className="mb-3"),

            # Second Row of Filters
            dbc.Row([
                dbc.Col([
                    html.Label('Category:', className="fw-bold"),
                    dcc.Dropdown(
                        id='fiction-filter',
                        options=[{'label': 'Fiction', 'value': 'Yes'}, {'label': 'Non-Fiction', 'value': 'No'}],
                        value=['Yes', 'No'],
                        multi=True,
                        placeholder="Select Category",
                        className="mb-2"
                    )
                ], md=3, sm=6, xs=12),
                dbc.Col([
                    html.Label('Media Type(s):', className="fw-bold"),
                    dcc.Dropdown(
                        id='media-filter',
                        options=[{'label': media, 'value': media} for media in sorted(data['Item Media'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Media Type(s)"
                    )
                ], md=4, sm=6, xs=12),
                dbc.Col([
                    html.Label('Select Publication Start Date:',
                        className="fw-bold",
                        style={'display': 'block'}
                    ),
                    dcc.DatePickerSingle(
                        id='publication-start-date-filter',
                        min_date_allowed=data['Title Publication Date'].min(),
                        max_date_allowed=data['Title Publication Date'].max(),
                        date=data['Title Publication Date'].min(),
                        display_format='YYYY-MM-DD',
                        className="mb-2"
                    ),
                ], md=2, sm=12, xs=12),
                dbc.Col([
                    html.Label('Select Publication End Date:',
                        className="fw-bold",
                        style={'display': 'block'}
                    ),
                    dcc.DatePickerSingle(
                        id='publication-end-date-filter',
                        min_date_allowed=data['Title Publication Date'].min(),
                        max_date_allowed=data['Title Publication Date'].max(),
                        date=data['Title Publication Date'].max(),
                        display_format='YYYY-MM-DD',
                    ),
                ], md=2, sm=12, xs=12)
            ], className="mb-3"),
        ]),
        className="mb-4",
        style={'boxShadow': '0 4px 8px rgba(0,0,0,0.1)', 'padding': '20px', 'borderRadius': '10px'}
    ),

    # KPIs
    dbc.Row([
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-book", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Unique Titles", className="card-title mt-2"),
                html.H3(id='total-titles', className="card-text"),
            ])
        ], color=colors['info'], inverse=True, className="text-center shadow-sm", style={'height': '165px', 'borderRadius': '10px'}), md=2, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-person", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Unique Authors", className="card-title mt-2"),
                html.H3(id='total-authors', className="card-text"),
            ])
        ], color=colors['success'], inverse=True, className="text-center shadow-sm", style={'height': '165px', 'borderRadius': '10px'}), md=2, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-people", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Unique Publishers", className="card-title mt-2"),
                html.H3(id='total-publishers', className="card-text"),
            ])
        ], color=colors['warning'], inverse=True, className="text-center shadow-sm", style={'height': '165px', 'borderRadius': '10px'}), md=2, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-calendar", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Earliest Publication Date", className="card-title mt-2"),
                html.H3(id='earliest-publication', className="card-text"),
            ])
        ], color=colors['primary'], inverse=True, className="text-center shadow-sm", style={'height': '165px', 'borderRadius': '10px'}), md=3, sm=6, xs=12),

        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.I(className="bi bi-calendar-check", style={'font-size': '2rem', 'color': 'white'}),
                html.H5("Latest Publication Date", className="card-title mt-2"),
                html.H3(id='latest-publication', className="card-text"),
            ])
        ], color=colors['secondary'], inverse=True, className="text-center shadow-sm", style={'height': '165px', 'borderRadius': '10px'}), md=3, sm=6, xs=12),
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
                    md=6,
                    sm=12,
                    className="mb-4"
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-category-distribution-donut',
                        type='default',
                        children=dcc.Graph(id='category-distribution-donut')
                    ),
                    md=6,
                    sm=12,
                    className="mb-4"
                ),
            ]),

            # Second Row: OverDrive Distribution Treemap
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-overdrive-distribution',
                        type='default',
                        children=dcc.Graph(id='overdrive-distribution')
                    ),
                    md=12,
                    sm=12,
                    className="mb-4"
                ),
            ]),

            # Third Row: Top 10 Publishers and Top 10 Authors Bar Charts
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-top-publishers',
                        type='default',
                        children=dcc.Graph(id='top-publishers-bar')
                    ),
                    md=6,
                    sm=12,
                    className="mb-4"
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-top-authors',
                        type='default',
                        children=dcc.Graph(id='top-authors-bar')
                    ),
                    md=6,
                    sm=12,
                    className="mb-4"
                ),
            ]),

            # Fourth Row: Publication Year Stacked Bar and Custom Chart
            dbc.Row([
                dbc.Col(
                    dcc.Loading(
                        id='loading-publication-year-stacked-bar',
                        type='default',
                        children=dcc.Graph(id='publication-year-stacked-bar')
                    ),
                    md=6,
                    sm=12,
                    className="mb-4"
                ),
                dbc.Col(
                    dcc.Loading(
                        id='loading-custom-chart',
                        type='default',
                        children=dcc.Graph(id='custom-chart')
                    ),
                    md=6,
                    sm=12,
                    className="mb-4"
                ),
            ]),
        ]),

        # Rank Trend Analysis Tab
        dbc.Tab(label='Rank Trend Analysis', tab_id='tab-detailed', children=[
            # First Row: Top Authors Slider
            dbc.Row([
                dbc.Col([
                    html.Label('Select Top N Authors:', className="fw-bold"),
                    dcc.Slider(
                        id='top-authors-slider',
                        min=5,
                        max=15,
                        step=1,
                        value=10,
                        marks={i: str(i) for i in range(5, 16, 1)},
                        tooltip={"placement": "bottom", "always_visible": True},
                        className="mb-4"
                    ),
                ], md=12, sm=12, xs=12),
            ], className="mb-4"),

            # Second Row: Heatmaps for Each Year
            dbc.Row([
                dbc.Col([
                    dcc.Loading(
                        id=f'loading-author-heatmap-{year}',
                        type='default',
                        children=dcc.Graph(id=f'author-heatmap-{year}')
                    )
                ], md=3, sm=6, xs=12) for year in HEATMAP_YEARS
            ], className="mb-4"),

            # Explanation for Average Rank of Top Authors
            dbc.Row([
                dbc.Col([
                    html.P(
                        "The heatmaps above display the average rank of the top authors across the years. "
                        "The average rank is calculated by taking the mean rank of each author's book titles in that year. "
                        "A lower rank indicates better popularity.",
                        className="text-muted"
                    )
                ], md=12, sm=12, xs=12),
            ], className="mb-2"),

            # Rank Trend Line Chart
            dbc.Row([
                dbc.Col([
                    html.Label('Select Title(s):', className="fw-bold"),
                    dcc.Dropdown(
                        id='title-filter',
                        options=[{'label': title, 'value': title} for title in sorted(data['Title Native Name'].unique())],
                        value=[],
                        multi=True,
                        placeholder="Select Title(s)",
                        className="mb-4"
                    ),
                    dcc.Loading(
                        id='loading-rank-trend-line',
                        type='default',
                        children=dcc.Graph(id='rank-trend-line')
                    ),
                ], md=12, sm=12, xs=12),
            ]),
        ]),
    ], id='tabs', active_tab='tab-overview'),

    # Footer
    html.Footer(
        dbc.Container([
            html.Hr(),
            html.P("© 2024 | Dashboard by Clarence Sai", className="text-center mb-0"),
        ], fluid=True),
        style={'padding': '20px 0', 'backgroundColor': colors['light']}
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
        treemap_data = filtered_data.groupby(['Title Publisher', 'Title Author']).size().reset_index(name='Count')

        fig = px.treemap(
            treemap_data,
            path=['Title Publisher', 'Title Author'],
            values='Count',
            title='OverDrive Distribution',
            color='Count',
            color_continuous_scale='Viridis',
            hover_data={'Count': True},
            labels={'Title Publisher': 'Publisher', 'Title Author': 'Author'}
        )
        fig.update_traces(
            hovertemplate='<b>Publisher: %{parent}</b><br><b>Author: %{label}</b><br>Count: %{value}<extra></extra>',
        )
        fig.update_layout(
            margin=dict(t=50, l=25, r=25, b=25),
            title={
                'text': "OverDrive Distribution",
                'y':0.95,
                'x':0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': {'size': 20}
            }
        )
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
        
        # Update the x-axis to show only the relevant years without sub-units
        fig.update_xaxes(
            tickmode='linear',
            tick0=HEATMAP_YEARS[0],
            dtick=1,
            range=[HEATMAP_YEARS[0] - 0.5, HEATMAP_YEARS[-1] + 0.5],  # Slightly beyond to ensure full bars are visible
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

def create_author_heatmap(filtered_data, year, top_n_authors=10):
    """
    Creates a heatmap for average ranks of top N authors in a specific year.
    
    Parameters:
    - filtered_data (DataFrame): The filtered dataset based on user selections.
    - year (int): The specific year for which the heatmap is created.
    - top_n_authors (int): Number of top authors to include in the heatmap.
    
    Returns:
    - Figure: A Plotly Heatmap figure.
    """
    # Filter data for the specific year
    year_data = filtered_data[filtered_data['Txn Calendar Year'] == year]
    
    if not year_data.empty:
        # Determine top N authors for this year based on the number of titles
        top_authors = year_data['Title Author'].value_counts().nlargest(top_n_authors).index.tolist()
        
        # Filter the data for these top authors
        year_data = year_data[year_data['Title Author'].isin(top_authors)]
        
        if not year_data.empty:
            # Aggregate data by 'Title Author'
            pivot_table = year_data.groupby('Title Author').agg(
                Average_Rank=('Rank', 'mean'),
                Ranks=('Rank', list),
                Titles=('Title Native Name', list)
            ).reset_index()
            
            # Sort authors by average rank (ascending)
            pivot_table_sorted = pivot_table.sort_values(by='Average_Rank')
            
            # Create the tooltip string with truncated titles
            pivot_table_sorted['Ranks_str'] = pivot_table_sorted.apply(
                lambda row: generate_ranks_str(row['Titles'], row['Ranks']),
                axis=1
            )
            
            # Reshape 'Ranks_str' to a 2D list for 'text' parameter
            text_reshaped = pivot_table_sorted['Ranks_str'].apply(lambda x: [x]).tolist()
            
            # Create Heatmap using go.Heatmap with 'text'
            fig = go.Figure(
                data=go.Heatmap(
                    z=pivot_table_sorted['Average_Rank'].values.reshape(-1, 1),  # Single column
                    x=['Average Rank'],  # Single label
                    y=pivot_table_sorted['Title Author'],
                    colorscale='Viridis',      # Use standard Viridis
                    reversescale=True,        # Lower ranks should have distinct color
                    colorbar=dict(title="Avg Rank"),
                    hoverongaps=False,
                    zmin=pivot_table_sorted['Average_Rank'].min(),
                    zmax=pivot_table_sorted['Average_Rank'].max(),
                    showscale=True,
                    text=text_reshaped,  # Use 'text' for hover information
                    hovertemplate=
                        '<b>%{y}</b><br>' +
                        'Average Rank: %{z:.1f}<br>' +
                        'Title-Rank:<br>' +
                        '%{text}<br>' +
                        '<extra></extra>',
                    hoverlabel=dict(
                        align='left',          
                        bgcolor="rgba(255, 255, 255, 0.9)",  # Semi-transparent white background
                        font=dict(
                            size=12,           
                            color="black",     
                            family="Arial"     
                        )
                    )
                )
            )
            
            # Update layout with increased top margin and normal y-axis orientation
            fig.update_layout(
                title=f'Average Rank of Top {top_n_authors} Authors in {year}',
                xaxis_title='',
                yaxis_title='Author',
                yaxis=dict(autorange='reversed'),  # Highest rank at top
                xaxis=dict(showticklabels=False),  # Hide x-axis labels
                margin=dict(t=150, l=200, r=50, b=50),  # Increased top margin
                template='plotly_white',  # Use a white template for better contrast
                hovermode='closest'  # Ensures that hover events are accurately captured
            )
            
            return fig
        else:
            # No data available for the top N authors
            fig = go.Figure()
            fig.update_layout(
                title=f"Average Rank of Top {top_n_authors} Authors in {year}",
                annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)],
                xaxis=dict(visible=False),
                yaxis=dict(visible=False),
                template='plotly_white'
            )
            return fig
    else:
        # No data available for the specific year
        fig = go.Figure()
        fig.update_layout(
            title=f"Average Rank of Top {top_n_authors} Authors in {year}",
            annotations=[dict(text="No data available", x=0.5, y=0.5, showarrow=False)],
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
            template='plotly_white'
        )
        return fig

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
                        tick0=HEATMAP_YEARS[0],
                        dtick=1,
                        range=[HEATMAP_YEARS[0], HEATMAP_YEARS[-1]],  # Adjusted to dynamic range
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
        Output('earliest-publication', 'children'),
        Output('latest-publication', 'children'),
        Output('media-type-donut', 'figure'),
        Output('category-distribution-donut', 'figure'),
        Output('overdrive-distribution', 'figure'),
        Output('top-publishers-bar', 'figure'),
        Output('top-authors-bar', 'figure'),
        Output('publication-year-stacked-bar', 'figure'),
        Output('custom-chart', 'figure'),
        # Outputs for each year's heatmap
        Output('author-heatmap-2020', 'figure'),
        Output('author-heatmap-2021', 'figure'),
        Output('author-heatmap-2022', 'figure'),
        Output('author-heatmap-2023', 'figure'),
        Output('rank-trend-line', 'figure'),
    ],
    [
        Input('year-filter', 'value'),
        Input('subject-filter', 'value'),
        Input('media-filter', 'value'),
        Input('publication-start-date-filter', 'date'),
        Input('publication-end-date-filter', 'date'),
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
    fig_overdrive_distribution = create_overdrive_distribution_treemap(filtered_data)
    fig_top_publishers = create_top_publishers_bar_chart(filtered_data)
    fig_top_authors = create_top_authors_bar_chart(filtered_data)
    fig_publication_year_stacked_bar = create_publication_year_stacked_bar_chart(filtered_data)
    fig_custom_chart = create_transaction_year_media_type_chart(filtered_data)

    # Create heatmaps for each year using Top N Authors
    heatmap_figs = []
    for year in HEATMAP_YEARS:
        fig = create_author_heatmap(filtered_data, year, top_n_authors)
        heatmap_figs.append(fig)

    # Create Rank Trend Line Chart
    fig_rank_trend = create_rank_trend_line_chart(filtered_data, selected_titles)

    return (
        total_titles,           
        total_authors,         
        total_publishers,       
        earliest_publication,   
        latest_publication,     
        fig_media_donut,
        fig_category_donut,
        fig_overdrive_distribution,
        fig_top_publishers,
        fig_top_authors,
        fig_publication_year_stacked_bar,
        fig_custom_chart,
        # Heatmap figures
        heatmap_figs[0] if len(heatmap_figs) > 0 else go.Figure(),
        heatmap_figs[1] if len(heatmap_figs) > 1 else go.Figure(),
        heatmap_figs[2] if len(heatmap_figs) > 2 else go.Figure(),
        heatmap_figs[3] if len(heatmap_figs) > 3 else go.Figure(),
        fig_rank_trend
    )

# Run the App
if __name__ == '__main__':
    app.run_server(debug=True)
