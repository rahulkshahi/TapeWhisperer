import dash
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import base64
import io
import json
from datetime import datetime
import re

# Initialize Dash app with Bootstrap theme
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Define validation rules
class ValidationRules:
    @staticmethod
    def validate_loan_number(value, index, all_values):
        """Validate and correct loan number"""
        if pd.isna(value) or not re.match(r'^[A-Z0-9]{6,}$', str(value)):
            return f"LN{datetime.now().strftime('%m%d')}{str(index).zfill(3)}"
        # Check uniqueness
        if all_values.count(value) > 1:
            return f"{value}_{index}"
        return value
    
    @staticmethod
    def validate_fico_score(value):
        """Validate and correct FICO score"""
        try:
            score = int(float(value))
            if score < 550:
                return 550
            elif score > 850:
                return 850
            return score
        except (ValueError, TypeError):
            return 650  # Default middle value
    
    @staticmethod
    def validate_house_val(value):
        """Validate and correct house value"""
        try:
            # Remove currency symbols and commas
            clean_value = str(value).replace('$', '').replace(',', '')
            num = float(clean_value)
            if num < 500:
                return 500
            elif num > 5000000:
                return 5000000
            return int(num)
        except (ValueError, TypeError):
            return 500000  # Default middle value
    
    @staticmethod
    def validate_kcltv(value):
        """Validate and correct kCLTV"""
        try:
            num = float(value)
            if num < 0.2:
                return 0.2
            elif num > 0.8:
                return 0.8
            return round(num, 3)
        except (ValueError, TypeError):
            return 0.5  # Default middle value
    
    @staticmethod
    def calculate_document_id(house_val, kcltv):
        """Calculate DocumentId based on business rules"""
        if house_val < 1000000 and 0.2 <= kcltv <= 0.3:
            return "21"
        elif 1000000 <= house_val < 4000000 and 0.3 < kcltv <= 0.5:
            return "22"
        else:
            return "27"

# Define the app layout
app.layout = dbc.Container([
    # Header
    dbc.Row([
        dbc.Col([
            html.H1("🔍 Excel Validation & Auto-Correction System", 
                   className="text-center text-primary mb-2"),
            html.P("Rule-based validation with intelligent data correction", 
                  className="text-center text-muted mb-4")
        ])
    ]),
    
    # Main content with two columns
    dbc.Row([
        # Left Column - Ruleset Generator
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(html.H4("📋 Step 1: Ruleset Generator", className="text-primary")),
                dbc.CardBody([
                    html.P("Generate standalone validation rulesets", className="text-muted"),
                    
                    # Rules Table
                    html.Div([
                        dbc.Table([
                            html.Thead([
                                html.Tr([
                                    html.Th("Field"),
                                    html.Th("Rule Type"),
                                    html.Th("Constraint"),
                                    html.Th("Auto-Correct")
                                ])
                            ]),
                            html.Tbody([
                                html.Tr([
                                    html.Td("LoanNumber"),
                                    html.Td("Format"),
                                    html.Td("Alphanumeric, Unique"),
                                    html.Td("Generate if invalid", className="text-info")
                                ]),
                                html.Tr([
                                    html.Td("FicoScore"),
                                    html.Td("Range"),
                                    html.Td("550-850"),
                                    html.Td("Clamp to boundary", className="text-info")
                                ]),
                                html.Tr([
                                    html.Td("HouseVal"),
                                    html.Td("Range"),
                                    html.Td("$500-$5M"),
                                    html.Td("Clamp to boundary", className="text-info")
                                ]),
                                html.Tr([
                                    html.Td("kCLTV"),
                                    html.Td("Range"),
                                    html.Td("0.2-0.8"),
                                    html.Td("Clamp to boundary", className="text-info")
                                ]),
                                html.Tr([
                                    html.Td("DocumentId"),
                                    html.Td("Conditional"),
                                    html.Td("21/22/27"),
                                    html.Td("Recalculate", className="text-info")
                                ])
                            ])
                        ], striped=True, bordered=True, hover=True, size="sm")
                    ]),
                    
                    # DocumentId Rules Alert
                    dbc.Alert([
                        html.H6("DocumentId Rules:", className="alert-heading"),
                        html.Ul([
                            html.Li("21: HouseVal < $1M AND kCLTV: 0.2-0.3"),
                            html.Li("22: HouseVal: $1M-$4M AND kCLTV: 0.3-0.5"),
                            html.Li("27: All other combinations")
                        ])
                    ], color="info", className="mt-3"),
                    
                    # Download buttons
                    html.Div([
                        html.H5("📥 Download Rulesets:", className="mt-3 mb-2"),
                        dbc.ButtonGroup([
                            dbc.Button("Download CSV Ruleset", 
                                     id="download-csv-btn", 
                                     color="success", 
                                     className="me-2"),
                            dbc.Button("Download JSON Rules", 
                                     id="download-json-btn", 
                                     color="secondary")
                        ])
                    ]),
                    dcc.Download(id="download-ruleset")
                ])
            ], className="mb-4")
        ], md=6),
        
        # Right Column - Sample Data Generator
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(html.H4("📊 Step 2: Generate Sample Data", className="text-primary")),
                dbc.CardBody([
                    html.P("Create sample Excel data for testing", className="text-muted"),
                    
                    dbc.InputGroup([
                        dbc.InputGroupText("Number of Rows:"),
                        dbc.Input(id="num-rows-input", type="number", value=10, min=1, max=1000)
                    ], className="mb-3"),
                    
                    dbc.Button("Generate Sample Data", 
                             id="generate-sample-btn", 
                             color="primary", 
                             className="w-100 mb-3"),
                    
                    # Sample data preview
                    html.Div(id="sample-data-preview"),
                    
                    # Download sample data
                    html.Div([
                        html.H5("📥 Download Sample:", className="mt-3 mb-2"),
                        dbc.Button("Download Sample Excel", 
                                 id="download-sample-btn", 
                                 color="success",
                                 disabled=True,
                                 className="w-100")
                    ]),
                    dcc.Download(id="download-sample")
                ])
            ], className="mb-4")
        ], md=6)
    ]),
    
    # Validation Section
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardHeader(html.H4("✅ Step 3: Validate & Correct Excel Data", className="text-primary")),
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            # File Upload Section
                            dbc.Card([
                                dbc.CardBody([
                                    html.H5("📤 Upload File:", className="text-info mb-3"),
                                    dcc.Upload(
                                        id='upload-data',
                                        children=html.Div([
                                            'Drag and Drop or ',
                                            html.A('Select Excel/CSV File', className="text-primary")
                                        ]),
                                        style={
                                            'width': '100%',
                                            'height': '100px',
                                            'lineHeight': '100px',
                                            'borderWidth': '2px',
                                            'borderStyle': 'dashed',
                                            'borderRadius': '10px',
                                            'textAlign': 'center',
                                            'backgroundColor': '#f8f9fa'
                                        },
                                        multiple=False,
                                        accept='.xlsx,.xls,.csv'
                                    ),
                                    html.Div(id='upload-status', className="mt-2")
                                ])
                            ], className="mb-3", color="light")
                        ], md=6),
                        
                        dbc.Col([
                            # Validation Actions
                            dbc.Card([
                                dbc.CardBody([
                                    html.H5("🔍 Validation Actions:", className="text-info mb-3"),
                                    dbc.Button("Validate & Correct Data", 
                                             id="validate-btn", 
                                             color="danger", 
                                             disabled=True,
                                             className="w-100 mb-2"),
                                    dbc.Button("Download Corrected Excel", 
                                             id="download-corrected-btn", 
                                             color="success",
                                             disabled=True,
                                             className="w-100")
                                ])
                            ], className="mb-3", color="light")
                        ], md=6)
                    ]),
                    
                    # Validation Results
                    html.Div(id="validation-results"),
                    dcc.Download(id="download-corrected"),
                    
                    # Store components for data
                    dcc.Store(id='stored-data'),
                    dcc.Store(id='corrected-data'),
                    dcc.Store(id='sample-data')
                ])
            ])
        ])
    ]),
    
    # Footer
    dbc.Row([
        dbc.Col([
            html.Hr(),
            html.P("© 2024 Excel Validation System | Built with Dash", 
                  className="text-center text-muted")
        ])
    ])
], fluid=True, className="p-4")

# Callbacks

@app.callback(
    [Output('download-ruleset', 'data'),
     Output('download-ruleset', 'filename')],
    [Input('download-csv-btn', 'n_clicks'),
     Input('download-json-btn', 'n_clicks')],
    prevent_initial_call=True
)
def download_ruleset(csv_clicks, json_clicks):
    """Generate and download ruleset files"""
    ctx = dash.callback_context
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'download-csv-btn':
        # Generate CSV ruleset
        ruleset_data = {
            'Field': ['LoanNumber', 'FicoScore', 'HouseVal', 'kCLTV', 'DocumentId'],
            'RuleType': ['format', 'range', 'range', 'range', 'conditional'],
            'MinValue': ['', '550', '500', '0.2', ''],
            'MaxValue': ['', '850', '5000000', '0.8', ''],
            'Pattern': ['^[A-Z0-9]{6,}$', '', '', '', ''],
            'ValidValues': ['', '', '', '', '21,22,27'],
            'AutoCorrectStrategy': [
                'Generate new if invalid/duplicate',
                'Clamp to nearest boundary',
                'Clamp to nearest boundary',
                'Clamp to nearest boundary',
                'Recalculate based on HouseVal & kCLTV'
            ]
        }
        df = pd.DataFrame(ruleset_data)
        return dcc.send_data_frame(df.to_csv, "validation_ruleset.csv", index=False), "validation_ruleset.csv"
    
    elif button_id == 'download-json-btn':
        # Generate JSON ruleset
        json_rules = {
            "version": "1.0",
            "rules": {
                "LoanNumber": {
                    "type": "format",
                    "pattern": "^[A-Z0-9]{6,}$",
                    "unique": True
                },
                "FicoScore": {
                    "type": "range",
                    "min": 550,
                    "max": 850,
                    "dataType": "integer"
                },
                "HouseVal": {
                    "type": "range",
                    "min": 500,
                    "max": 5000000,
                    "dataType": "number"
                },
                "kCLTV": {
                    "type": "range",
                    "min": 0.2,
                    "max": 0.8,
                    "dataType": "decimal"
                },
                "DocumentId": {
                    "type": "conditional",
                    "validValues": ["21", "22", "27"]
                }
            },
            "documentIdLogic": {
                "21": "HouseVal < 1000000 AND kCLTV between 0.2-0.3",
                "22": "HouseVal 1000000-4000000 AND kCLTV between 0.3-0.5",
                "27": "All other cases"
            }
        }
        return dcc.send_string(json.dumps(json_rules, indent=2), "validation_rules.json"), "validation_rules.json"
    
    return None, None

@app.callback(
    [Output('sample-data', 'data'),
     Output('sample-data-preview', 'children'),
     Output('download-sample-btn', 'disabled')],
    [Input('generate-sample-btn', 'n_clicks')],
    [State('num-rows-input', 'value')],
    prevent_initial_call=True
)
def generate_sample_data(n_clicks, num_rows):
    """Generate sample data"""
    if n_clicks:
        # Generate sample data
        data = []
        for i in range(num_rows):
            # Create diverse scenarios
            scenario = i % 4
            
            # Generate loan number
            loan_number = f"LN{datetime.now().strftime('%Y')}{str(i+1).zfill(4)}"
            
            # Generate FICO score
            fico_score = np.random.randint(550, 851)
            
            # Generate house value and kCLTV based on scenario
            if scenario == 0:  # DocumentId 21 scenario
                house_val = np.random.randint(100000, 999999)
                kcltv = round(np.random.uniform(0.2, 0.3), 3)
            elif scenario == 1:  # DocumentId 22 scenario
                house_val = np.random.randint(1000000, 3999999)
                kcltv = round(np.random.uniform(0.31, 0.5), 3)
            elif scenario == 2:  # DocumentId 27 scenario (high value)
                house_val = np.random.randint(4000000, 5000000)
                kcltv = round(np.random.uniform(0.5, 0.8), 3)
            else:  # Random scenario
                house_val = np.random.randint(500, 5000000)
                kcltv = round(np.random.uniform(0.2, 0.8), 3)
            
            # Calculate DocumentId
            doc_id = ValidationRules.calculate_document_id(house_val, kcltv)
            
            data.append({
                'LoanNumber': loan_number,
                'FicoScore': fico_score,
                'HouseVal': house_val,
                'kCLTV': kcltv,
                'DocumentId': doc_id
            })
        
        df = pd.DataFrame(data)
        
        # Create preview table
        preview = dbc.Table.from_dataframe(
            df.head(5), 
            striped=True, 
            bordered=True, 
            hover=True,
            responsive=True,
            size="sm"
        )
        
        preview_component = html.Div([
            dbc.Alert(f"Generated {num_rows} rows of sample data", color="success"),
            html.H6("Preview (first 5 rows):"),
            preview
        ])
        
        return df.to_dict('records'), preview_component, False
    
    return None, None, True

@app.callback(
    Output('download-sample', 'data'),
    Input('download-sample-btn', 'n_clicks'),
    State('sample-data', 'data'),
    prevent_initial_call=True
)
def download_sample_excel(n_clicks, data):
    """Download sample data as Excel"""
    if n_clicks and data:
        df = pd.DataFrame(data)
        return dcc.send_data_frame(df.to_excel, "sample_loan_data.xlsx", index=False)
    return None

@app.callback(
    [Output('stored-data', 'data'),
     Output('upload-status', 'children'),
     Output('validate-btn', 'disabled')],
    Input('upload-data', 'contents'),
    State('upload-data', 'filename'),
    prevent_initial_call=True
)
def upload_file(contents, filename):
    """Handle file upload"""
    if contents:
        try:
            content_type, content_string = contents.split(',')
            decoded = base64.b64decode(content_string)
            
            # Parse file based on extension
            if filename.endswith('.csv'):
                df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
            else:
                df = pd.read_excel(io.BytesIO(decoded))
            
            status = dbc.Alert([
                html.I(className="bi bi-check-circle-fill me-2"),
                f"Successfully uploaded: {filename}",
                html.Br(),
                f"Rows: {len(df)}, Columns: {len(df.columns)}"
            ], color="success")
            
            return df.to_dict('records'), status, False
            
        except Exception as e:
            status = dbc.Alert([
                html.I(className="bi bi-x-circle-fill me-2"),
                f"Error uploading file: {str(e)}"
            ], color="danger")
            return None, status, True
    
    return None, None, True

@app.callback(
    [Output('corrected-data', 'data'),
     Output('validation-results', 'children'),
     Output('download-corrected-btn', 'disabled')],
    Input('validate-btn', 'n_clicks'),
    State('stored-data', 'data'),
    prevent_initial_call=True
)
def validate_and_correct(n_clicks, data):
    """Validate and correct the uploaded data"""
    if n_clicks and data:
        df = pd.DataFrame(data)
        corrections_log = []
        
        # Track corrections
        corrected_df = df.copy()
        correction_flags = pd.DataFrame(False, index=df.index, columns=df.columns)
        
        # Validate LoanNumber
        loan_numbers = df['LoanNumber'].tolist()
        for idx, value in enumerate(loan_numbers):
            new_value = ValidationRules.validate_loan_number(value, idx, loan_numbers)
            if str(value) != str(new_value):
                corrected_df.loc[idx, 'LoanNumber'] = new_value
                correction_flags.loc[idx, 'LoanNumber'] = True
                corrections_log.append(f"Row {idx+1}: LoanNumber '{value}' → '{new_value}'")
        
        # Validate FicoScore
        for idx, value in df['FicoScore'].items():
            new_value = ValidationRules.validate_fico_score(value)
            if str(value) != str(new_value):
                corrected_df.loc[idx, 'FicoScore'] = new_value
                correction_flags.loc[idx, 'FicoScore'] = True
                corrections_log.append(f"Row {idx+1}: FicoScore '{value}' → '{new_value}'")
        
        # Validate HouseVal
        for idx, value in df['HouseVal'].items():
            new_value = ValidationRules.validate_house_val(value)
            if str(value) != str(new_value):
                corrected_df.loc[idx, 'HouseVal'] = new_value
                correction_flags.loc[idx, 'HouseVal'] = True
                corrections_log.append(f"Row {idx+1}: HouseVal '{value}' → '{new_value}'")
        
        # Validate kCLTV
        for idx, value in df['kCLTV'].items():
            new_value = ValidationRules.validate_kcltv(value)
            if str(value) != str(new_value):
                corrected_df.loc[idx, 'kCLTV'] = new_value
                correction_flags.loc[idx, 'kCLTV'] = True
                corrections_log.append(f"Row {idx+1}: kCLTV '{value}' → '{new_value}'")
        
        # Recalculate DocumentId
        for idx in corrected_df.index:
            house_val = corrected_df.loc[idx, 'HouseVal']
            kcltv = corrected_df.loc[idx, 'kCLTV']
            new_doc_id = ValidationRules.calculate_document_id(house_val, kcltv)
            old_doc_id = df.loc[idx, 'DocumentId'] if 'DocumentId' in df.columns else None
            
            if str(old_doc_id) != str(new_doc_id):
                corrected_df.loc[idx, 'DocumentId'] = new_doc_id
                correction_flags.loc[idx, 'DocumentId'] = True
                corrections_log.append(f"Row {idx+1}: DocumentId '{old_doc_id}' → '{new_doc_id}'")
        
        # Calculate statistics
        total_rows = len(df)
        total_corrections = len(corrections_log)
        corrected_rows = correction_flags.any(axis=1).sum()
        valid_rows = total_rows - corrected_rows
        
        # Create results display
        results = html.Div([
            # Statistics cards
            dbc.Row([
                dbc.Col([
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Total Rows", className="text-muted"),
                            html.H2(str(total_rows), className="text-primary")
                        ])
                    ])
                ], md=3),
                dbc.Col([
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Valid Rows", className="text-muted"),
                            html.H2(str(valid_rows), className="text-success")
                        ])
                    ])
                ], md=3),
                dbc.Col([
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Corrected Rows", className="text-muted"),
                            html.H2(str(corrected_rows), className="text-warning")
                        ])
                    ])
                ], md=3),
                dbc.Col([
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Success Rate", className="text-muted"),
                            html.H2(f"{(valid_rows/total_rows*100):.1f}%", className="text-info")
                        ])
                    ])
                ], md=3)
            ], className="mb-3"),
            
            # Corrections log
            html.Div([
                html.H5("Corrections Log:", className="mt-3"),
                dbc.Alert([
                    html.Ul([html.Li(correction) for correction in corrections_log[:10]])
                    if corrections_log else "No corrections needed!"
                ], color="warning" if corrections_log else "success")
            ]) if len(corrections_log) <= 10 else html.Div([
                html.H5("Corrections Log:", className="mt-3"),
                dbc.Alert([
                    html.Ul([html.Li(correction) for correction in corrections_log[:10]]),
                    html.P(f"... and {len(corrections_log)-10} more corrections", className="text-muted")
                ], color="warning")
            ]),
            
            # Data table with highlighting
            html.H5("Corrected Data Preview:", className="mt-3"),
            html.Div([
                dash_table.DataTable(
                    data=corrected_df.head(10).to_dict('records'),
                    columns=[{'name': col, 'id': col} for col in corrected_df.columns],
                    style_cell={'textAlign': 'center'},
                    style_data_conditional=[
                        {
                            'if': {
                                'row_index': i,
                                'column_id': col
                            },
                            'backgroundColor': '#fff3cd',
                            'fontWeight': 'bold',
                            'fontStyle': 'italic'
                        }
                        for i in range(min(10, len(correction_flags)))
                        for col in correction_flags.columns
                        if i < len(correction_flags) and correction_flags.iloc[i][col]
                    ],
                    style_header={
                        'backgroundColor': 'rgb(230, 230, 230)',
                        'fontWeight': 'bold'
                    }
                )
            ])
        ])
        
        # Store corrected data and corrections for download
        corrected_data_with_flags = {
            'data': corrected_df.to_dict('records'),
            'corrections': correction_flags.to_dict('records')
        }
        
        return corrected_data_with_flags, results, False
    
    return None, None, True

@app.callback(
    Output('download-corrected', 'data'),
    Input('download-corrected-btn', 'n_clicks'),
    State('corrected-data', 'data'),
    prevent_initial_call=True
)
def download_corrected_excel(n_clicks, corrected_data):
    """Download corrected Excel file"""
    if n_clicks and corrected_data:
        df = pd.DataFrame(corrected_data['data'])
        
        # Create Excel file with multiple sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write corrected data
            df.to_excel(writer, sheet_name='Corrected Data', index=False)
            
            # Add formatting for corrected cells
            workbook = writer.book
            worksheet = writer.sheets['Corrected Data']
            
            # Define formats
            corrected_format = workbook.add_format({
                'italic': True,
                'bold': True,
                'bg_color': '#fff3cd'
            })
            
            # Apply formatting to corrected cells
            corrections = pd.DataFrame(corrected_data['corrections'])
            for row_idx in range(len(df)):
                for col_idx, col_name in enumerate(df.columns):
                    if row_idx < len(corrections) and corrections.iloc[row_idx][col_name]:
                        worksheet.write(row_idx + 1, col_idx, df.iloc[row_idx][col_name], corrected_format)
        
        output.seek(0)
        return dcc.send_bytes(output.read(), "../../../../../../../Downloads/validated_data.xlsx")
    
    return None

# Run the app
if __name__ == '__main__':
    app.run(debug=True)