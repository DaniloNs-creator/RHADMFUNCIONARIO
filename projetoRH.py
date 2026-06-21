import streamlit as st
import pandas as pd
import sqlite3
import datetime
import plotly.express as px
from streamlit_option_menu import option_menu
import base64
from io import BytesIO

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="XCM & Duatlo Coach Pro",
    page_icon="🚴",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
    /* MAIN THEME */
    .main {
        background: linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 100%);
        color: #f0f0f0;
    }
    .stApp {
        background: #0f0f1a;
    }
    /* CARDS */
    .card {
        background: rgba(30, 30, 50, 0.8);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 25px;
        margin: 15px 0;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 8px 32px rgba(0,0,0,0.4);
        transition: all 0.3s ease;
    }
    .card:hover {
        transform: translateY(-5px);
        border-color: #00d4ff;
        box-shadow: 0 12px 40px rgba(0,212,255,0.15);
    }
    .metric-card {
        background: rgba(0, 212, 255, 0.05);
        border-left: 4px solid #00d4ff;
        padding: 15px 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .title-gradient {
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(135deg, #00d4ff, #7b2ffc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        letter-spacing: -1px;
    }
    .section-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: #00d4ff;
        border-bottom: 2px solid rgba(0,212,255,0.2);
        padding-bottom: 8px;
        margin-top: 30px;
    }
    .highlight {
        color: #ff6b6b;
        font-weight: 600;
    }
    /* SIDEBAR */
    .css-1d391kg {
        background: rgba(15, 15, 26, 0.95);
        border-right: 1px solid rgba(255,255,255,0.05);
    }
    .stSelectbox, .stTextInput, .stNumberInput {
        background: rgba(255,255,255,0.05);
        border-radius: 10px;
        border: 1px solid rgba(255,255,255,0.1);
        color: white;
    }
    /* BUTTONS */
    .stButton > button {
        background: linear-gradient(135deg, #00d4ff, #7b2ffc);
        color: white;
        border: none;
        border-radius: 50px;
        padding: 12px 30px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0,212,255,0.3);
        width: 100%;
    }
    .stButton > button:hover {
        transform: scale(1.02);
        box-shadow: 0 8px 25px rgba(0,212,255,0.5);
    }
    /* TABLES */
    .dataframe {
        background: rgba(30, 30, 50, 0.6);
        border-radius: 15px;
        padding: 10px;
        border: 1px solid rgba(255,255,255,0.05);
    }
    /* PROGRESS BARS */
    .stProgress > div > div {
        background: linear-gradient(90deg, #00d4ff, #7b2ffc);
    }
    /* SCROLLBAR */
    ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
    }
    ::-webkit-scrollbar-track {
        background: rgba(255,255,255,0.05);
    }
    ::-webkit-scrollbar-thumb {
        background: #00d4ff;
        border-radius: 10px;
    }
    /* RESPONSIVE */
    @media (max-width: 768px) {
        .title-gradient {
            font-size: 2rem;
        }
        .card {
            padding: 15px;
        }
    }
</style>
""", unsafe_allow_html=True)

# ---------------- DATABASE ----------------
def init_db():
    conn = sqlite3.connect('coach_data.db')
    c = conn.cursor()
    # Workouts table
    c.execute('''CREATE TABLE IF NOT EXISTS workouts
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  date TEXT,
                  type TEXT,
                  duration_minutes REAL,
                  hr_avg INTEGER,
                  cadence_avg INTEGER,
                  distance_km REAL,
                  calories_burned INTEGER,
                  notes TEXT)''')
    # Nutrition table
    c.execute('''CREATE TABLE IF NOT EXISTS nutrition
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  date TEXT,
                  meal_type TEXT,
                  food_name TEXT,
                  quantity REAL,
                  unit TEXT,
                  calories REAL,
                  protein REAL,
                  carbs REAL,
                  fat REAL)''')
    # Body metrics table
    c.execute('''CREATE TABLE IF NOT EXISTS body_metrics
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  date TEXT,
                  weight_kg REAL,
                  body_fat REAL,
                  muscle_mass REAL,
                  notes TEXT)''')
    # Food database
    c.execute('''CREATE TABLE IF NOT EXISTS food_db
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  food_name TEXT UNIQUE,
                  calories_per_100g REAL,
                  protein_per_100g REAL,
                  carbs_per_100g REAL,
                  fat_per_100g REAL)''')
    
    # Insert default foods if empty
    c.execute("SELECT COUNT(*) FROM food_db")
    if c.fetchone()[0] == 0:
        default_foods = [
            ('Peito de Frango', 165, 31, 0, 3.6),
            ('Arroz Branco', 130, 2.7, 28, 0.3),
            ('Batata Doce', 86, 1.6, 20, 0.1),
            ('Ovo', 155, 13, 1.1, 11),
            ('Aveia', 389, 16.9, 66, 6.9),
            ('Banana', 89, 1.1, 23, 0.3),
            ('Whey Protein', 350, 80, 5, 3),
            ('Salada Mista', 15, 1.5, 3, 0.1),
            ('Azeite de Oliva', 884, 0, 0, 100),
            ('Pão Integral', 265, 9, 45, 3.5),
            ('Queijo Branco', 98, 21, 2, 1.5),
            ('Iogurte Grego', 59, 10, 3.5, 0.2),
        ]
        for food in default_foods:
            c.execute("INSERT OR IGNORE INTO food_db VALUES (NULL, ?, ?, ?, ?, ?)", food)
    
    conn.commit()
    conn.close()

init_db()

# ---------------- USER PROFILE ----------------
USER_PROFILE = {
    'age': 29,
    'height': 1.90,
    'weight': 110,
    'target_weight': 78,
    'target_date': '2026-12-31',
    'fc_max': 178,
    'resting_hr': 60,  # estimated
    'level': 'intermediate',
    'disponibility_weekday': 1.5,  # hours
    'disponibility_weekend': 5  # hours
}

# Calculate HR Zones (based on Karvonen)
def calculate_hr_zones(fc_max, resting_hr=60):
    hr_reserve = fc_max - resting_hr
    zones = {
        'Z1 (Recovery)': (resting_hr + 0.5 * hr_reserve, resting_hr + 0.6 * hr_reserve),
        'Z2 (Endurance)': (resting_hr + 0.6 * hr_reserve, resting_hr + 0.7 * hr_reserve),
        'Z3 (Tempo)': (resting_hr + 0.7 * hr_reserve, resting_hr + 0.8 * hr_reserve),
        'Z4 (Threshold)': (resting_hr + 0.8 * hr_reserve, resting_hr + 0.9 * hr_reserve),
        'Z5 (VO2max)': (resting_hr + 0.9 * hr_reserve, fc_max)
    }
    return zones

HR_ZONES = calculate_hr_zones(USER_PROFILE['fc_max'])

# ---------------- CALORIE CALCULATIONS ----------------
def calculate_bmr(weight, height, age):
    # Mifflin-St Jeor
    return 10 * weight + 6.25 * height * 100 - 5 * age + 5

def calculate_tdee(weight, height, age, activity_factor=1.6):
    bmr = calculate_bmr(weight, height, age)
    return bmr * activity_factor

# ---------------- NUTRITION PLAN ----------------
def generate_nutrition_plan(weight, height, age, target_weight, target_date):
    bmr = calculate_bmr(weight, height, age)
    
    # Activity factor for high training load (1.7)
    tdee = bmr * 1.7
    
    # Calculate deficit needed
    current_date = datetime.datetime.now()
    target_date_obj = datetime.datetime.strptime(target_date, '%Y-%m-%d')
    days_remaining = (target_date_obj - current_date).days
    
    total_weight_loss = weight - target_weight  # 32kg
    calories_per_kg_fat = 7700
    total_calorie_deficit = total_weight_loss * calories_per_kg_fat
    daily_deficit = total_calorie_deficit / days_remaining if days_remaining > 0 else 1000
    
    # Cap deficit at 1000 to avoid undernourishment
    daily_deficit = min(daily_deficit, 1000)
    
    # Ensure minimum intake
    min_intake = 1800
    target_calories = max(tdee - daily_deficit, min_intake)
    
    # Macro split (40% Carbs, 30% Protein, 30% Fat)
    carbs = (target_calories * 0.40) / 4
    protein = (target_calories * 0.30) / 4
    fat = (target_calories * 0.30) / 9
    
    # Adjust protein for athlete (2.0g/kg)
    protein_athlete = weight * 2.0
    protein = max(protein, protein_athlete)
    
    return {
        'bmr': bmr,
        'tdee': tdee,
        'target_calories': target_calories,
        'daily_deficit': daily_deficit,
        'carbs_g': carbs,
        'protein_g': protein,
        'fat_g': fat,
        'days_remaining': days_remaining
    }

# ---------------- TRAINING PLAN ----------------
def generate_training_plan(week_number):
    # Periodization: 4-week blocks
    block = (week_number - 1) // 4 + 1
    week_in_block = (week_number - 1) % 4 + 1
    
    # Intensity progression
    if block == 1:  # Foundation
        volume_mult = 0.7 + (week_in_block - 1) * 0.1
        intensity = 'Low-Moderate (Z1-Z2)'
        bike_hr_zones = ['Z2', 'Z2', 'Z2']
    elif block == 2:  # Build
        volume_mult = 0.9 + (week_in_block - 1) * 0.05
        intensity = 'Moderate-High (Z2-Z3)'
        bike_hr_zones = ['Z2', 'Z3', 'Z2']
    elif block == 3:  # Intensity
        volume_mult = 0.95 + (week_in_block - 1) * 0.05
        intensity = 'High (Z3-Z4)'
        bike_hr_zones = ['Z3', 'Z4', 'Z3']
    else:  # Peak
        volume_mult = 0.8 + (week_in_block - 1) * 0.1
        intensity = 'Very High (Z4-Z5)'
        bike_hr_zones = ['Z3', 'Z4', 'Z4']
    
    # Weekly schedule (days 1-7)
    schedule = {
        1: {'type': 'Musculação', 'focus': 'Força', 'duration': 1.5},
        2: {'type': 'MTB + Musculação', 'focus': 'Endurance + Força', 'duration': 1.5},
        3: {'type': 'Corrida + Musculação', 'focus': 'Condicionamento', 'duration': 1.5},
        4: {'type': 'MTB + Musculação', 'focus': 'Resistência', 'duration': 1.5},
        5: {'type': 'Corrida + Musculação', 'focus': 'Ritmo', 'duration': 1.5},
        6: {'type': 'MTB Longo', 'focus': 'Resistência', 'duration': 4.5},
        7: {'type': 'Corrida Longa', 'focus': 'Resistência', 'duration': 4.0}
    }
    
    # Bike specific workouts
    bike_workouts = {
        2: {'duration': 60, 'zones': bike_hr_zones[0], 'cadence': 85},
        4: {'duration': 70, 'zones': bike_hr_zones[1], 'cadence': 90},
        6: {'duration': 180, 'zones': bike_hr_zones[2], 'cadence': 80},
    }
    
    return {
        'week': week_number,
        'block': block,
        'volume_mult': volume_mult,
        'intensity': intensity,
        'schedule': schedule,
        'bike_workouts': bike_workouts
    }

# ---------------- MAIN APP ----------------
def main():
    # Sidebar
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 20px 0;">
            <span style="font-size: 3rem;">🚴</span>
            <h2 style="color: #00d4ff;">XCM Coach</h2>
            <p style="color: #888; font-size: 0.9rem;">Duatlo & MTB Expert</p>
        </div>
        """, unsafe_allow_html=True)
        
        selected = option_menu(
            menu_title=None,
            options=["Dashboard", "Treinos", "Nutrição", "Progresso", "Coach AI"],
            icons=["house", "bicycle", "egg", "graph-up", "robot"],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background": "transparent"},
                "icon": {"color": "#00d4ff", "font-size": "1.2rem"},
                "nav-link": {
                    "color": "#888",
                    "font-size": "1rem",
                    "margin": "5px 0",
                    "border-radius": "10px",
                    "padding": "12px 15px",
                    "transition": "all 0.3s"
                },
                "nav-link-selected": {
                    "background": "rgba(0, 212, 255, 0.15)",
                    "color": "#00d4ff",
                    "border-left": "3px solid #00d4ff"
                }
            }
        )
        
        st.markdown("---")
        st.markdown("""
        <div style="font-size: 0.8rem; color: #555; padding: 10px;">
            <p>⚡ <span style="color: #888;">Peso meta: 78kg</span></p>
            <p>📅 <span style="color: #888;">Meta: 31/12/2026</span></p>
            <p>❤️ <span style="color: #888;">FC Máx: 178 bpm</span></p>
        </div>
        """, unsafe_allow_html=True)

    # Main content
    if selected == "Dashboard":
        render_dashboard()
    elif selected == "Treinos":
        render_training()
    elif selected == "Nutrição":
        render_nutrition()
    elif selected == "Progresso":
        render_progress()
    elif selected == "Coach AI":
        render_coach()

def render_dashboard():
    st.markdown('<div class="title-gradient">🏆 Dashboard do Atleta</div>', unsafe_allow_html=True)
    
    # Nutrition plan
    plan = generate_nutrition_plan(
        USER_PROFILE['weight'],
        USER_PROFILE['height'],
        USER_PROFILE['age'],
        USER_PROFILE['target_weight'],
        USER_PROFILE['target_date']
    )
    
    # Current week
    start_date = datetime.datetime(2026, 1, 1)
    today = datetime.datetime.now()
    week_number = ((today - start_date).days // 7) + 1
    week_number = max(1, min(week_number, 26))  # 6 months
    
    training = generate_training_plan(week_number)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="card">
            <h4 style="color: #00d4ff; margin-bottom: 15px;">📊 Status Atual</h4>
            <div class="metric-card">
                <span style="color: #888;">Peso</span>
                <h2 style="color: white;">{USER_PROFILE['weight']} kg</h2>
                <span style="color: #ff6b6b;">Meta: {USER_PROFILE['target_weight']} kg</span>
            </div>
            <div class="metric-card">
                <span style="color: #888;">IMC</span>
                <h2 style="color: white;">{USER_PROFILE['weight'] / (USER_PROFILE['height']**2):.1f}</h2>
            </div>
            <div class="metric-card">
                <span style="color: #888;">Semana</span>
                <h2 style="color: white;">{week_number} / 26</h2>
                <span style="color: #00d4ff;">Bloco {training['block']}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        deficit = plan['daily_deficit']
        pct_progress = min(100, (110 - USER_PROFILE['weight']) / (110 - 78) * 100)
        
        st.markdown(f"""
        <div class="card">
            <h4 style="color: #00d4ff; margin-bottom: 15px;">🔥 Déficit Calórico</h4>
            <div class="metric-card">
                <span style="color: #888;">Déficit Diário</span>
                <h2 style="color: #ff6b6b;">-{deficit:.0f} kcal</h2>
            </div>
            <div class="metric-card">
                <span style="color: #888;">Meta Diária</span>
                <h2 style="color: white;">{plan['target_calories']:.0f} kcal</h2>
            </div>
            <div style="margin-top: 15px;">
                <span style="color: #888;">Progresso emagrecimento</span>
                <div style="background: rgba(255,255,255,0.1); border-radius: 10px; height: 10px; margin: 8px 0;">
                    <div style="background: linear-gradient(90deg, #00d4ff, #7b2ffc); width: {pct_progress}%; height: 100%; border-radius: 10px;"></div>
                </div>
                <span style="color: #00d4ff;">{pct_progress:.1f}%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="card">
            <h4 style="color: #00d4ff; margin-bottom: 15px;">📅 Treino de Hoje</h4>
            <div style="background: rgba(0,212,255,0.05); border-radius: 15px; padding: 15px; border: 1px solid rgba(0,212,255,0.1);">
                <p style="font-size: 1.2rem; font-weight: 600; color: white;">{training['schedule'][today.isoweekday()]['type']}</p>
                <p style="color: #888;">Foco: {training['schedule'][today.isoweekday()]['focus']}</p>
                <p style="color: #888;">Duração: {training['schedule'][today.isoweekday()]['duration']}h</p>
                <p style="color: #00d4ff;">Intensidade: {training['intensity']}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # HR Zones
    st.markdown('<div class="section-title">❤️ Zonas de Frequência Cardíaca</div>', unsafe_allow_html=True)
    cols = st.columns(5)
    for i, (zone, (low, high)) in enumerate(HR_ZONES.items()):
        with cols[i]:
            st.markdown(f"""
            <div class="card" style="text-align: center; padding: 15px;">
                <p style="color: #888; font-size: 0.8rem;">{zone}</p>
                <h3 style="color: white;">{int(low)}-{int(high)}</h3>
                <p style="color: #555; font-size: 0.7rem;">bpm</p>
            </div>
            """, unsafe_allow_html=True)

def render_training():
    st.markdown('<div class="title-gradient">🚴 Plano de Treinos</div>', unsafe_allow_html=True)
    
    start_date = datetime.datetime(2026, 1, 1)
    today = datetime.datetime.now()
    week_number = ((today - start_date).days // 7) + 1
    week_number = max(1, min(week_number, 26))
    
    training = generate_training_plan(week_number)
    
    st.markdown(f"""
    <div class="card">
        <h4 style="color: #00d4ff;">Semana {week_number} - Bloco {training['block']}</h4>
        <p><span style="color: #888;">Intensidade:</span> {training['intensity']}</p>
        <p><span style="color: #888;">Volume:</span> {training['volume_mult']:.0%} da carga máxima</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Weekly schedule
    days = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
    cols = st.columns(7)
    
    for i, day in enumerate(days):
        day_num = i + 1
        workout = training['schedule'].get(day_num, {})
        with cols[i]:
            st.markdown(f"""
            <div class="card" style="padding: 12px; min-height: 180px;">
                <h5 style="color: #00d4ff; font-size: 0.9rem;">{day}</h5>
                <p style="font-size: 0.8rem; color: white;">{workout.get('type', 'Descanso')}</p>
                <p style="font-size: 0.7rem; color: #888;">{workout.get('focus', '')}</p>
                <p style="font-size: 0.7rem; color: #555;">{workout.get('duration', 0)}h</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Bike specific workouts
    st.markdown('<div class="section-title">🚵 Treinos de MTB Detalhados</div>', unsafe_allow_html=True)
    
    for day_num, bike_wk in training['bike_workouts'].items():
        zones = bike_wk['zones'].split('+')
        hr_range = []
        for z in zones:
            if z in HR_ZONES:
                low, high = HR_ZONES[z]
                hr_range.append(f"{int(low)}-{int(high)}")
        
        st.markdown(f"""
        <div class="card">
            <h5 style="color: #00d4ff;">{days[day_num-1]}</h5>
            <p><span style="color: #888;">Duração:</span> {bike_wk['duration']} min</p>
            <p><span style="color: #888;">Zona FC:</span> {bike_wk['zones']} ({' / '.join(hr_range)} bpm)</p>
            <p><span style="color: #888;">Cadência Meta:</span> {bike_wk['cadence']} RPM</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Log workout
    st.markdown('<div class="section-title">📝 Log de Treino</div>', unsafe_allow_html=True)
    
    with st.form("workout_log"):
        col1, col2 = st.columns(2)
        with col1:
            workout_type = st.selectbox("Tipo", ["MTB", "Corrida", "Musculação", "Duatlo"])
            duration = st.number_input("Duração (minutos)", min_value=10, max_value=300, value=60)
            hr_avg = st.number_input("FC Média (bpm)", min_value=60, max_value=200, value=140)
        with col2:
            cadence = st.number_input("Cadência Média (RPM)", min_value=40, max_value=120, value=85) if workout_type in ["MTB", "Duatlo"] else 0
            distance = st.number_input("Distância (km)", min_value=0.0, max_value=200.0, value=20.0)
            notes = st.text_area("Observações")
        
        if st.form_submit_button("Salvar Treino"):
            conn = sqlite3.connect('coach_data.db')
            c = conn.cursor()
            c.execute("""INSERT INTO workouts 
                         (date, type, duration_minutes, hr_avg, cadence_avg, distance_km, notes)
                         VALUES (?, ?, ?, ?, ?, ?, ?)""",
                      (datetime.datetime.now().strftime('%Y-%m-%d'),
                       workout_type, duration, hr_avg, cadence, distance, notes))
            conn.commit()
            conn.close()
            st.success("✅ Treino registrado com sucesso!")

def render_nutrition():
    st.markdown('<div class="title-gradient">🍽️ Plano Nutricional</div>', unsafe_allow_html=True)
    
    plan = generate_nutrition_plan(
        USER_PROFILE['weight'],
        USER_PROFILE['height'],
        USER_PROFILE['age'],
        USER_PROFILE['target_weight'],
        USER_PROFILE['target_date']
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        <div class="card">
            <h4 style="color: #00d4ff;">Metas Diárias</h4>
            <div class="metric-card">
                <span style="color: #888;">Calorias</span>
                <h2 style="color: white;">{plan['target_calories']:.0f} kcal</h2>
            </div>
            <div class="metric-card">
                <span style="color: #888;">Proteínas</span>
                <h2 style="color: #ff6b6b;">{plan['protein_g']:.0f} g</h2>
                <span style="color: #888;">{plan['protein_g']*4:.0f} kcal</span>
            </div>
            <div class="metric-card">
                <span style="color: #888;">Carboidratos</span>
                <h2 style="color: #00d4ff;">{plan['carbs_g']:.0f} g</h2>
                <span style="color: #888;">{plan['carbs_g']*4:.0f} kcal</span>
            </div>
            <div class="metric-card">
                <span style="color: #888;">Gorduras</span>
                <h2 style="color: #ffd93d;">{plan['fat_g']:.0f} g</h2>
                <span style="color: #888;">{plan['fat_g']*9:.0f} kcal</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="card">
            <h4 style="color: #00d4ff;">Estratégia</h4>
            <p><span style="color: #888;">TMB:</span> {plan['bmr']:.0f} kcal</p>
            <p><span style="color: #888;">Gasto Total:</span> {plan['tdee']:.0f} kcal</p>
            <p><span style="color: #ff6b6b;">Déficit:</span> -{plan['daily_deficit']:.0f} kcal/dia</p>
            <p><span style="color: #888;">Dias restantes:</span> {plan['days_remaining']}</p>
            <div style="margin-top: 15px; background: rgba(0,212,255,0.05); padding: 10px; border-radius: 10px;">
                <p style="color: #888; font-size: 0.9rem;">💡 Distribuição: 40% Carbs • 30% Protein • 30% Fat</p>
                <p style="color: #888; font-size: 0.9rem;">🍗 2.0g de proteína/kg de peso para atleta</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # Food log
    st.markdown('<div class="section-title">📝 Log Alimentar</div>', unsafe_allow_html=True)
    
    conn = sqlite3.connect('coach_data.db')
    
    with st.form("food_log"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            meal_type = st.selectbox("Refeição", ["Café da Manhã", "Lanche Manhã", "Almoço", "Lanche Tarde", "Jantar", "Ceia"])
            food_name = st.text_input("Alimento", placeholder="Ex: Peito de Frango")
        
        with col2:
            quantity = st.number_input("Quantidade", min_value=0.0, value=100.0)
            unit = st.selectbox("Unidade", ["g", "ml", "unidade", "colher"])
        
        with col3:
            # Auto-fill from food database
            conn_db = sqlite3.connect('coach_data.db')
            foods_df = pd.read_sql_query("SELECT food_name, calories_per_100g, protein_per_100g, carbs_per_100g, fat_per_100g FROM food_db", conn_db)
            conn_db.close()
            
            if not foods_df.empty:
                food_options = foods_df['food_name'].tolist()
                selected_food = st.selectbox("Ou escolha do banco", [""] + food_options)
                if selected_food:
                    food_data = foods_df[foods_df['food_name'] == selected_food].iloc[0]
                    calories = food_data['calories_per_100g'] * (quantity / 100)
                    protein = food_data['protein_per_100g'] * (quantity / 100)
                    carbs = food_data['carbs_per_100g'] * (quantity / 100)
                    fat = food_data['fat_per_100g'] * (quantity / 100)
                    st.info(f"💡 {selected_food}: {calories:.0f} kcal | P:{protein:.1f}g C:{carbs:.1f}g G:{fat:.1f}g")
                else:
                    calories = st.number_input("Calorias", min_value=0.0, value=100.0)
                    protein = st.number_input("Proteína (g)", min_value=0.0, value=10.0)
                    carbs = st.number_input("Carboidratos (g)", min_value=0.0, value=15.0)
                    fat = st.number_input("Gordura (g)", min_value=0.0, value=5.0)
            else:
                calories = st.number_input("Calorias", min_value=0.0, value=100.0)
                protein = st.number_input("Proteína (g)", min_value=0.0, value=10.0)
                carbs = st.number_input("Carboidratos (g)", min_value=0.0, value=15.0)
                fat = st.number_input("Gordura (g)", min_value=0.0, value=5.0)
        
        if st.form_submit_button("Salvar Refeição"):
            conn_insert = sqlite3.connect('coach_data.db')
            c = conn_insert.cursor()
            c.execute("""INSERT INTO nutrition 
                         (date, meal_type, food_name, quantity, unit, calories, protein, carbs, fat)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                      (datetime.datetime.now().strftime('%Y-%m-%d'),
                       meal_type, food_name, quantity, unit, calories, protein, carbs, fat))
            conn_insert.commit()
            conn_insert.close()
            st.success("✅ Refeição registrada!")
    
    # Show today's log
    today_str = datetime.datetime.now().strftime('%Y-%m-%d')
    today_nutrition = pd.read_sql_query(
        f"SELECT meal_type, food_name, quantity, unit, calories, protein, carbs, fat FROM nutrition WHERE date='{today_str}'",
        conn
    )
    conn.close()
    
    if not today_nutrition.empty:
        st.markdown(f"### 📊 Resumo de Hoje - {today_str}")
        
        total_cal = today_nutrition['calories'].sum()
        total_protein = today_nutrition['protein'].sum()
        total_carbs = today_nutrition['carbs'].sum()
        total_fat = today_nutrition['fat'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("🔥 Calorias", f"{total_cal:.0f} kcal", f"{total_cal - plan['target_calories']:.0f} do plano")
        col2.metric("💪 Proteína", f"{total_protein:.1f}g", f"{total_protein - plan['protein_g']:.1f}g")
        col3.metric("🌾 Carboidratos", f"{total_carbs:.1f}g", f"{total_carbs - plan['carbs_g']:.1f}g")
        col4.metric("🧈 Gordura", f"{total_fat:.1f}g", f"{total_fat - plan['fat_g']:.1f}g")
        
        st.dataframe(today_nutrition[['meal_type', 'food_name', 'quantity', 'unit', 'calories']], use_container_width=True)

def render_progress():
    st.markdown('<div class="title-gradient">📈 Progresso</div>', unsafe_allow_html=True)
    
    conn = sqlite3.connect('coach_data.db')
    
    # Body metrics
    body_df = pd.read_sql_query("SELECT date, weight_kg FROM body_metrics ORDER BY date", conn)
    
    # Workout stats
    workout_df = pd.read_sql_query("SELECT date, type, duration_minutes, distance_km FROM workouts ORDER BY date", conn)
    
    # Nutrition stats
    nutrition_df = pd.read_sql_query("SELECT date, calories FROM nutrition ORDER BY date", conn)
    daily_calories = nutrition_df.groupby('date')['calories'].sum().reset_index()
    
    conn.close()
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("📊 Evolução do Peso")
        
        if not body_df.empty:
            fig = px.line(body_df, x='date', y='weight_kg', 
                         title='Evolução do Peso',
                         labels={'weight_kg': 'Peso (kg)', 'date': 'Data'})
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font_color='white', xaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
                            yaxis=dict(gridcolor='rgba(255,255,255,0.1)'))
            fig.add_hline(y=78, line_dash="dash", line_color="#ff6b6b", 
                         annotation_text="Meta 78kg")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Ainda não há dados de peso registrados.")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("🚴 Volume de Treino")
        
        if not workout_df.empty:
            weekly_volume = workout_df.groupby(pd.to_datetime(workout_df['date']).dt.isocalendar().week).agg({
                'duration_minutes': 'sum',
                'distance_km': 'sum'
            }).reset_index()
            
            fig = px.bar(weekly_volume, x='week', y='duration_minutes',
                        title='Volume Semanal de Treino (min)',
                        labels={'duration_minutes': 'Minutos', 'week': 'Semana'})
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font_color='white', xaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
                            yaxis=dict(gridcolor='rgba(255,255,255,0.1)'))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Ainda não há dados de treino registrados.")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Nutrition summary
    st.markdown('<div class="section-title">📊 Resumo Nutricional</div>', unsafe_allow_html=True)
    
    if not daily_calories.empty:
        fig = px.line(daily_calories, x='date', y='calories',
                     title='Ingestão Calórica Diária',
                     labels={'calories': 'Calorias', 'date': 'Data'})
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                         font_color='white', xaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
                         yaxis=dict(gridcolor='rgba(255,255,255,0.1)'))
        
        plan = generate_nutrition_plan(
            USER_PROFILE['weight'],
            USER_PROFILE['height'],
            USER_PROFILE['age'],
            USER_PROFILE['target_weight'],
            USER_PROFILE['target_date']
        )
        fig.add_hline(y=plan['target_calories'], line_dash="dash", line_color="#00d4ff",
                     annotation_text="Meta Diária")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Ainda não há dados nutricionais registrados.")
    
    # Log weight
    st.markdown('<div class="section-title">⚖️ Registrar Peso</div>', unsafe_allow_html=True)
    
    with st.form("weight_log"):
        weight = st.number_input("Peso (kg)", min_value=50.0, max_value=200.0, value=float(USER_PROFILE['weight']))
        body_fat = st.number_input("% Gordura (opcional)", min_value=0.0, max_value=50.0, value=25.0)
        notes = st.text_area("Observações")
        
        if st.form_submit_button("Salvar Peso"):
            conn = sqlite3.connect('coach_data.db')
            c = conn.cursor()
            c.execute("""INSERT INTO body_metrics 
                         (date, weight_kg, body_fat, notes)
                         VALUES (?, ?, ?, ?)""",
                      (datetime.datetime.now().strftime('%Y-%m-%d'),
                       weight, body_fat, notes))
            conn.commit()
            conn.close()
            USER_PROFILE['weight'] = weight
            st.success("✅ Peso registrado com sucesso!")

def render_coach():
    st.markdown('<div class="title-gradient">🤖 Coach IA</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="card">
        <h4 style="color: #00d4ff;">💬 Assistente Virtual</h4>
        <p style="color: #888;">Faça perguntas sobre treino, nutrição ou sobre seu progresso.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Quick stats for context
    plan = generate_nutrition_plan(
        USER_PROFILE['weight'],
        USER_PROFILE['height'],
        USER_PROFILE['age'],
        USER_PROFILE['target_weight'],
        USER_PROFILE['target_date']
    )
    
    conn = sqlite3.connect('coach_data.db')
    last_workout = pd.read_sql_query("SELECT date, type, duration_minutes, distance_km FROM workouts ORDER BY date DESC LIMIT 1", conn)
    conn.close()
    
    st.markdown("""
    <div class="card" style="background: rgba(0,212,255,0.05); border: 1px solid rgba(0,212,255,0.2);">
        <p style="color: #00d4ff;">📌 Contexto atual:</p>
        <ul style="color: #888; list-style: none; padding: 0;">
            <li>• Peso: <span style="color: white;">{} kg</span> (Meta: 78kg)</li>
            <li>• Déficit diário: <span style="color: #ff6b6b;">-{} kcal</span></li>
            <li>• Semana de treino: <span style="color: white;">{}ª semana</span></li>
            <li>• Último treino: {}</li>
        </ul>
    </div>
    """.format(
        USER_PROFILE['weight'],
        int(plan['daily_deficit']),
        ((datetime.datetime.now() - datetime.datetime(2026, 1, 1)).days // 7) + 1,
        f"{last_workout.iloc[0]['type']} ({last_workout.iloc[0]['duration_minutes']}min)" if not last_workout.empty else "Nenhum registrado"
    ), unsafe_allow_html=True)
    
    # Chat interface
    st.markdown("### 💬 Pergunte ao Coach")
    user_question = st.text_area("Sua pergunta:", placeholder="Ex: Como ajustar minha dieta nos dias de treino longo?")
    
    if st.button("Enviar Pergunta", use_container_width=True):
        if user_question:
            # Generate response based on keywords
            response = generate_coach_response(user_question, plan)
            st.markdown(f"""
            <div class="card" style="background: rgba(123, 47, 252, 0.1); border: 1px solid rgba(123, 47, 252, 0.2);">
                <p style="color: #00d4ff;">🤖 Coach:</p>
                <p style="color: white;">{response}</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.warning("Por favor, digite uma pergunta.")

def generate_coach_response(question, plan):
    """Simple rule-based coach responses"""
    question_lower = question.lower()
    
    if any(word in question_lower for word in ['dieta', 'comer', 'alimento', 'nutrição']):
        return f"""
        Com base no seu perfil, sua ingestão calórica diária deve ser de {plan['target_calories']:.0f} kcal.
        Distribuição: {plan['carbs_g']:.0f}g de carboidratos, {plan['protein_g']:.0f}g de proteína e {plan['fat_g']:.0f}g de gordura.
        
        💡 Dica: Em dias de treino longo, aumente os carboidratos para 4-5g/kg de peso (440-550g). Em dias de descanso, reduza para 2-3g/kg e aumente a proteína.
        """
    
    elif any(word in question_lower for word in ['treino', 'bike', 'mtb', 'corrida']):
        return """
        Sua semana de treino deve ser periodizada:
        • Segunda: Musculação (Força)
        • Terça: MTB 60min (Z2) + Musculação
        • Quarta: Corrida 45min + Musculação
        • Quinta: MTB 70min (Z3) + Musculação
        • Sexta: Corrida 50min + Musculação
        • Sábado: MTB Longo (3h Z2)
        • Domingo: Corrida Longa (2h Z2)
        
        💡 Lembre-se: Aumente a intensidade gradualmente, e mantenha a cadência do MTB entre 80-90 RPM.
        """
    
    elif any(word in question_lower for word in ['peso', 'emagrecimento', 'perder']):
        return f"""
        Você está em um processo de emagrecimento com déficit de {plan['daily_deficit']:.0f} kcal/dia.
        Para chegar em 78kg até 31/12/2026, faltam {plan['days_remaining']} dias.
        
        🔑 Chave: Mantenha o déficit, mas nunca abaixo de 1800kcal para sustentar seus treinos.
        Priorize alimentos proteicos (2g/kg) e carboidratos complexos antes dos treinos.
        """
    
    elif any(word in question_lower for word in ['cardio', 'fc', 'frequência', 'coração']):
        hr_zones = calculate_hr_zones(USER_PROFILE['fc_max'])
        response = "Suas Zonas de FC (baseado em 178bpm):\n"
        for zone, (low, high) in hr_zones.items():
            response += f"• {zone}: {int(low)}-{int(high)} bpm\n"
        return response
    
    else:
        return """
        😊 Aqui estão algumas recomendações gerais para você:
        
        1. **Hidratação**: Beba 3-4L de água por dia, especialmente nos dias de treino.
        2. **Recuperação**: Durma 7-9h por noite. O sono é crucial para perda de peso e performance.
        3. **Treino**: Respeite as Zonas de FC - Z2 para treinos de resistência, Z3 para tempo, Z4 para limiar.
        4. **Nutrição**: Priorize alimentos integrais, verduras e proteínas magras.
        
        Para perguntas específicas, seja mais detalhado sobre o que precisa.
        """

# ---------------- RUN APP ----------------
if __name__ == "__main__":
    main()