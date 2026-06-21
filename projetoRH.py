# app.py
import streamlit as st
import sqlite3
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import json
from pathlib import Path
import calendar
from dataclasses import dataclass
from typing import List, Dict, Optional
import hashlib

# Configuração da página
st.set_page_config(
    page_title="PRO CYCLIST AI | Coach Premium",
    page_icon="🚴‍♂️",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://www.seuapp.com/help',
        'Report a bug': "https://www.seuapp.com/bug",
        'About': "# Pro Cyclist AI Coach v2.0\nTreinamento de Elite"
    }
)

# CSS Customizado Profissional
def load_css():
    st.markdown("""
    <style>
    /* Design System */
    :root {
        --primary: #FF4B4B;
        --primary-dark: #CC0000;
        --secondary: #0068C9;
        --success: #00C853;
        --warning: #FFD700;
        --danger: #FF1744;
        --dark: #0E1117;
        --darker: #000000;
        --light: #FAFAFA;
        --gray: #808495;
        --card-bg: #1E1E1E;
        --border: #2E2E2E;
    }
    
    /* Global Styles */
    .stApp {
        background: linear-gradient(135deg, #0a0a0a 0%, #1a1a2e 100%);
    }
    
    /* Hero Section */
    .hero-section {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 3rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        color: white;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    }
    
    .hero-title {
        font-size: 3.5rem;
        font-weight: 800;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .hero-subtitle {
        font-size: 1.5rem;
        opacity: 0.95;
        margin-top: 1rem;
    }
    
    /* Metric Cards */
    .metric-card {
        background: linear-gradient(145deg, #1e1e2f 0%, #2a2a3f 100%);
        border-radius: 20px;
        padding: 2rem;
        border: 1px solid #3a3a5f;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        cursor: pointer;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 40px rgba(102, 126, 234, 0.4);
        border-color: #667eea;
    }
    
    .metric-value {
        font-size: 2.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #a0a0b0;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    
    .metric-delta {
        font-size: 0.9rem;
        color: #00C853;
        font-weight: 600;
    }
    
    /* Progress Bar Custom */
    .custom-progress {
        background: #1e1e2f;
        border-radius: 20px;
        padding: 1.5rem;
        margin: 1.5rem 0;
        border: 1px solid #3a3a5f;
    }
    
    .progress-fill {
        height: 30px;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        transition: width 0.6s ease;
        position: relative;
        overflow: hidden;
    }
    
    .progress-fill::after {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
        animation: shimmer 2s infinite;
    }
    
    @keyframes shimmer {
        0% { transform: translateX(-100%); }
        100% { transform: translateX(100%); }
    }
    
    /* Food Cards */
    .food-card {
        background: linear-gradient(145deg, #1e1e2f 0%, #252540 100%);
        border-radius: 15px;
        padding: 1.5rem;
        border: 1px solid #3a3a5f;
        margin: 1rem 0;
        transition: all 0.3s ease;
    }
    
    .food-card:hover {
        border-color: #667eea;
        transform: scale(1.02);
    }
    
    /* Workout Cards */
    .workout-card {
        background: linear-gradient(145deg, #1e1e2f 0%, #252540 100%);
        border-radius: 15px;
        padding: 1.5rem;
        border-left: 5px solid #667eea;
        margin: 1rem 0;
        transition: all 0.3s ease;
    }
    
    .workout-card.musculacao {
        border-left-color: #FF4B4B;
    }
    
    .workout-card.ciclismo {
        border-left-color: #0068C9;
    }
    
    .workout-card.corrida {
        border-left-color: #00C853;
    }
    
    /* Zone Badges */
    .zone-badge {
        display: inline-block;
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.85rem;
        margin: 0.2rem;
    }
    
    .zone-z1 { background: #00C853; color: white; }
    .zone-z2 { background: #64DD17; color: black; }
    .zone-z3 { background: #FFD700; color: black; }
    .zone-z4 { background: #FF9100; color: white; }
    .zone-z5 { background: #FF1744; color: white; }
    
    /* Animations */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .animate-fade-in {
        animation: fadeInUp 0.6s ease-out;
    }
    
    /* Timeline */
    .timeline {
        position: relative;
        padding: 2rem 0;
    }
    
    .timeline::before {
        content: '';
        position: absolute;
        left: 50%;
        top: 0;
        bottom: 0;
        width: 2px;
        background: linear-gradient(to bottom, #667eea, #764ba2);
        transform: translateX(-50%);
    }
    
    .timeline-item {
        margin: 2rem 0;
        padding: 1.5rem;
        background: linear-gradient(145deg, #1e1e2f 0%, #252540 100%);
        border-radius: 15px;
        border: 1px solid #3a3a5f;
        position: relative;
    }
    
    .timeline-item::before {
        content: '';
        position: absolute;
        width: 20px;
        height: 20px;
        background: #667eea;
        border-radius: 50%;
        left: -10px;
        top: 50%;
        transform: translateY(-50%);
    }
    
    /* Buttons */
    .btn-primary {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem 2rem;
        border-radius: 10px;
        border: none;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.3s ease;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .btn-primary:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.5);
    }
    
    /* Responsive Grid */
    .responsive-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 1.5rem;
        padding: 1rem 0;
    }
    
    /* Glassmorphism */
    .glass-card {
        background: rgba(30, 30, 47, 0.7);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 2rem;
    }
    
    /* Scrollbar */
    ::-webkit-scrollbar {
        width: 10px;
    }
    
    ::-webkit-scrollbar-track {
        background: #0a0a0a;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea, #764ba2);
        border-radius: 5px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #764ba2, #667eea);
    }
    </style>
    """, unsafe_allow_html=True)

# Classe de Gerenciamento de Dados
@dataclass
class AthleteProfile:
    name: str = "Atleta Pro"
    age: int = 29
    height: float = 1.90
    current_weight: float = 110.0
    target_weight: float = 78.0
    fcm: int = 178
    target_date: str = "2026-12-31"
    
class DatabaseManager:
    def __init__(self):
        self.conn = sqlite3.connect('pro_cyclist_ai.db', check_same_thread=False)
        self.create_tables()
        self.initialize_food_database()
        
    def create_tables(self):
        c = self.conn.cursor()
        
        c.executescript('''
            CREATE TABLE IF NOT EXISTS athlete_profile (
                id INTEGER PRIMARY KEY,
                name TEXT,
                age INTEGER,
                height REAL,
                start_weight REAL,
                target_weight REAL,
                fcm INTEGER,
                start_date TEXT,
                target_date TEXT
            );
            
            CREATE TABLE IF NOT EXISTS workouts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                type TEXT NOT NULL,
                subtype TEXT,
                duration_min INTEGER,
                distance_km REAL,
                avg_hr INTEGER,
                max_hr INTEGER,
                avg_cadence INTEGER,
                zone TEXT,
                rpe INTEGER,
                calories_burned INTEGER,
                feeling TEXT,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            CREATE TABLE IF NOT EXISTS meals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                meal_type TEXT NOT NULL,
                food_name TEXT NOT NULL,
                portion_g REAL,
                calories REAL,
                protein_g REAL,
                carbs_g REAL,
                fat_g REAL,
                fiber_g REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            CREATE TABLE IF NOT EXISTS weight_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                weight_kg REAL,
                body_fat_pct REAL,
                notes TEXT
            );
            
            CREATE TABLE IF NOT EXISTS food_database (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                category TEXT,
                calories_per_100g REAL,
                protein_per_100g REAL,
                carbs_per_100g REAL,
                fat_per_100g REAL,
                fiber_per_100g REAL
            );
            
            CREATE TABLE IF NOT EXISTS weekly_plan (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_number INTEGER,
                day_of_week TEXT,
                workout_type TEXT,
                description TEXT,
                duration_planned INTEGER,
                target_zone TEXT,
                target_cadence TEXT,
                rpe_target INTEGER
            );
        ''')
        
        self.conn.commit()
    
    def initialize_food_database(self):
        c = self.conn.cursor()
        
        # Verifica se já existem dados
        c.execute("SELECT COUNT(*) FROM food_database")
        if c.fetchone()[0] > 0:
            return
            
        # Banco de dados completo de alimentos
        foods = [
            # Proteínas
            ("Peito de Frango Grelhado", "Proteínas", 165, 31, 0, 3.6, 0),
            ("Salmão", "Proteínas", 208, 20, 0, 13, 0),
            ("Atum em Água", "Proteínas", 116, 26, 0, 1, 0),
            ("Ovo Inteiro", "Proteínas", 155, 13, 1.1, 11, 0),
            ("Clara de Ovo", "Proteínas", 52, 11, 0.7, 0.2, 0),
            ("Carne Bovina Magra", "Proteínas", 250, 26, 0, 15, 0),
            ("Peito de Peru", "Proteínas", 135, 30, 0, 1, 0),
            ("Tilápia", "Proteínas", 96, 20, 0, 1.7, 0),
            ("Whey Protein", "Suplementos", 380, 80, 10, 5, 0),
            ("Caseína", "Suplementos", 360, 80, 5, 2, 0),
            ("BCAA", "Suplementos", 0, 0, 0, 0, 0),
            ("Creatina", "Suplementos", 0, 0, 0, 0, 0),
            
            # Carboidratos
            ("Arroz Integral Cozido", "Carboidratos", 123, 2.6, 25.6, 1, 1.8),
            ("Batata Doce Cozida", "Carboidratos", 86, 1.6, 20.1, 0.1, 3),
            ("Aveia em Flocos", "Carboidratos", 389, 16.9, 66.3, 6.9, 10.6),
            ("Pão Integral", "Carboidratos", 247, 13, 41, 3.4, 7),
            ("Macarrão Integral", "Carboidratos", 124, 5.3, 26.5, 0.5, 4.5),
            ("Quinoa Cozida", "Carboidratos", 120, 4.4, 21.3, 1.9, 2.8),
            ("Banana", "Frutas", 89, 1.1, 22.8, 0.3, 2.6),
            ("Maçã", "Frutas", 52, 0.3, 13.8, 0.2, 2.4),
            ("Mamão", "Frutas", 43, 0.5, 10.8, 0.3, 1.7),
            ("Morango", "Frutas", 32, 0.7, 7.7, 0.3, 2),
            
            # Gorduras Saudáveis
            ("Abacate", "Gorduras", 160, 2, 8.5, 14.7, 6.7),
            ("Azeite de Oliva", "Gorduras", 884, 0, 0, 100, 0),
            ("Castanha do Pará", "Gorduras", 656, 14.3, 12.3, 66.4, 7.5),
            ("Amêndoas", "Gorduras", 579, 21.2, 21.6, 49.9, 12.5),
            ("Pasta de Amendoim", "Gorduras", 588, 25, 20, 50, 6),
            ("Semente de Chia", "Gorduras", 486, 16.5, 42.1, 30.7, 34.4),
            ("Linhaça", "Gorduras", 534, 18.3, 28.9, 42.2, 27.3),
            
            # Vegetais
            ("Brócolis Cozido", "Vegetais", 35, 2.4, 7.2, 0.4, 3.3),
            ("Espinafre", "Vegetais", 23, 2.9, 3.6, 0.4, 2.2),
            ("Alface", "Vegetais", 15, 1.4, 2.9, 0.2, 1.3),
            ("Tomate", "Vegetais", 18, 0.9, 3.9, 0.2, 1.2),
            ("Cenoura Crua", "Vegetais", 41, 0.9, 9.6, 0.2, 2.8),
            ("Couve", "Vegetais", 49, 4.3, 8.8, 0.9, 3.6),
            
            # Laticínios
            ("Iogurte Grego Natural", "Laticínios", 97, 10, 3.6, 5, 0),
            ("Leite Desnatado", "Laticínios", 34, 3.4, 5, 0.1, 0),
            ("Queijo Cottage", "Laticínios", 98, 11.1, 3.4, 4.3, 0),
            ("Queijo Mussarela Light", "Laticínios", 250, 25, 3, 15, 0),
            
            # Bebidas
            ("Água de Coco", "Bebidas", 19, 0.7, 4.2, 0, 0),
            ("Suco de Laranja Natural", "Bebidas", 45, 0.7, 10.4, 0.2, 0.2),
            ("Café Preto", "Bebidas", 2, 0.1, 0, 0, 0),
            ("Chá Verde", "Bebidas", 1, 0.1, 0.2, 0, 0),
        ]
        
        c.executemany('''
            INSERT OR IGNORE INTO food_database 
            (name, category, calories_per_100g, protein_per_100g, carbs_per_100g, fat_per_100g, fiber_per_100g)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', foods)
        
        self.conn.commit()

# Calculadora de Nutrição
class NutritionCalculator:
    @staticmethod
    def calculate_bmr(weight: float, height: float, age: int, gender: str = "male") -> float:
        """Harris-Benedict Equation"""
        if gender == "male":
            return 88.362 + (13.397 * weight) + (4.799 * height * 100) - (5.677 * age)
        else:
            return 447.593 + (9.247 * weight) + (3.098 * height * 100) - (4.330 * age)
    
    @staticmethod
    def calculate_tdee(bmr: float, activity_level: str = "very_active") -> float:
        """Total Daily Energy Expenditure"""
        multipliers = {
            "sedentary": 1.2,
            "light": 1.375,
            "moderate": 1.55,
            "active": 1.725,
            "very_active": 1.9
        }
        return bmr * multipliers.get(activity_level, 1.725)
    
    @staticmethod
    def calculate_macros(weight: float, target_weight: float, months_to_goal: int) -> Dict:
        """Calculate optimal macros for weight loss while maintaining performance"""
        
        bmr = NutritionCalculator.calculate_bmr(weight, 1.90, 29)
        tdee = NutritionCalculator.calculate_tdee(bmr)
        
        # Déficit progressivo baseado no tempo até a meta
        weight_to_lose = weight - target_weight
        weeks = months_to_goal * 4.33
        deficit_per_week = (weight_to_lose * 7700) / weeks  # 7700 kcal = 1kg gordura
        
        # Limitar déficit diário para preservar performance
        daily_deficit = min(deficit_per_week / 7, 750)
        
        # Calorias alvo
        target_calories = tdee - daily_deficit
        
        # Macros para performance
        protein_g = weight * 2.0  # 2g/kg para preservação muscular
        fat_g = weight * 0.8  # 0.8g/kg para saúde hormonal
        carbs_g = (target_calories - (protein_g * 4) - (fat_g * 9)) / 4
        
        return {
            "bmr": round(bmr),
            "tdee": round(tdee),
            "target_calories": round(target_calories),
            "deficit": round(daily_deficit),
            "protein_g": round(protein_g),
            "carbs_g": round(carbs_g),
            "fat_g": round(fat_g),
            "protein_pct": round(protein_g * 4 / target_calories * 100),
            "carbs_pct": round(carbs_g * 4 / target_calories * 100),
            "fat_pct": round(fat_g * 9 / target_calories * 100),
            "weekly_loss_kg": round(deficit_per_week / 7700, 1)
        }

# Gerador de Planos de Treino
class WorkoutPlanner:
    @staticmethod
    def calculate_zones(fcm: int) -> Dict:
        return {
            "Z1": {"name": "Recuperação", "min": 0, "max": int(fcm * 0.65), "color": "#00C853"},
            "Z2": {"name": "Endurance", "min": int(fcm * 0.66), "max": int(fcm * 0.75), "color": "#64DD17"},
            "Z3": {"name": "Tempo", "min": int(fcm * 0.76), "max": int(fcm * 0.85), "color": "#FFD700"},
            "Z4": {"name": "Limiar", "min": int(fcm * 0.86), "max": int(fcm * 0.92), "color": "#FF9100"},
            "Z5": {"name": "VO2 Max", "min": int(fcm * 0.93), "max": fcm, "color": "#FF1744"}
        }
    
    @staticmethod
    def generate_weekly_plan(week: int, total_weeks: int = 26) -> Dict:
        """Gera plano semanal progressivo baseado na semana atual"""
        
        # Progressão de volume baseada na semana
        progress = min(week / total_weeks, 1.0)
        
        # Volume base (horas)
        base_volume = 8 + (progress * 7)  # 8h -> 15h
        
        weekly_plan = {
            "Monday": {
                "day": "Segunda-feira",
                "workouts": [
                    {
                        "type": "Musculação",
                        "focus": "Peito + Tríceps + Core",
                        "duration": 60,
                        "exercises": [
                            "Supino Reto: 4x10 (80% RM)",
                            "Crucifixo Inclinado: 3x12",
                            "Tríceps Corda: 4x12",
                            "Paralelas: 3x falha",
                            "Prancha: 4x60s",
                            "Abdominal Roda: 3x15"
                        ]
                    },
                    {
                        "type": "Ciclismo MTB",
                        "focus": "Endurance Z2",
                        "duration": 90,
                        "details": {
                            "warmup": "15min Z1 (90-100 rpm)",
                            "main": "60min Z2 (116-133 bpm, 85-95 rpm)",
                            "cooldown": "15min Z1 (90-100 rpm)"
                        }
                    }
                ]
            },
            "Tuesday": {
                "day": "Terça-feira",
                "workouts": [
                    {
                        "type": "Musculação",
                        "focus": "Costas + Bíceps",
                        "duration": 60,
                        "exercises": [
                            "Barra Fixa: 4x8-10",
                            "Remada Curvada: 4x10",
                            "Puxada Frente: 3x12",
                            "Remada Unilateral: 3x12",
                            "Rosca Direta: 3x12",
                            "Rosca Martelo: 3x12"
                        ]
                    },
                    {
                        "type": "Corrida",
                        "focus": "Técnica + Endurance Z2",
                        "duration": 45,
                        "details": {
                            "warmup": "10min Z1",
                            "main": "25min Z2 (116-133 bpm)",
                            "technique": "10min educativos (skipping, hopping)"
                        }
                    }
                ]
            },
            "Wednesday": {
                "day": "Quarta-feira",
                "workouts": [
                    {
                        "type": "Musculação",
                        "focus": "Pernas + Core",
                        "duration": 60,
                        "exercises": [
                            "Agachamento Livre: 5x8 (85% RM)",
                            "Leg Press: 4x12",
                            "Stiff: 4x10",
                            "Cadeira Extensora: 3x15",
                            "Panturrilha: 4x20",
                            "Prancha Lateral: 3x45s cada"
                        ]
                    },
                    {
                        "type": "Ciclismo MTB",
                        "focus": "Intervalado Z3/Z4",
                        "duration": 90,
                        "details": {
                            "warmup": "20min Z1-Z2",
                            "intervals": "5x (5min Z4 152-163 bpm + 3min Z2)",
                            "cooldown": "15min Z1",
                            "cadence": "80-95 rpm nos intervalos"
                        }
                    }
                ]
            },
            "Thursday": {
                "day": "Quinta-feira",
                "workouts": [
                    {
                        "type": "Musculação",
                        "focus": "Ombros + Abdômen",
                        "duration": 60,
                        "exercises": [
                            "Desenvolvimento Halteres: 4x10",
                            "Elevação Lateral: 4x15",
                            "Elevação Frontal: 3x12",
                            "Facepull: 3x15",
                            "Abdominal Infra: 4x20",
                            "Prancha Spiderman: 3x12 cada"
                        ]
                    },
                    {
                        "type": "Corrida",
                        "focus": "Fartlek Z1-Z3",
                        "duration": 45,
                        "details": {
                            "warmup": "10min Z1",
                            "fartlek": "8x (1min Z3 + 2min Z1)",
                            "cooldown": "10min Z1"
                        }
                    }
                ]
            },
            "Friday": {
                "day": "Sexta-feira",
                "workouts": [
                    {
                        "type": "Musculação",
                        "focus": "Full Body Power",
                        "duration": 60,
                        "exercises": [
                            "Levantamento Terra: 5x5 (85% RM)",
                            "Supino Inclinado: 4x8",
                            "Agachamento Frontal: 4x8",
                            "Remada Curvada: 4x8",
                            "Clean & Press: 3x6"
                        ]
                    },
                    {
                        "type": "Ciclismo MTB",
                        "focus": "Recuperação Ativa Z1",
                        "duration": 60,
                        "details": {
                            "zone": "Z1 (<116 bpm)",
                            "cadence": "90-100 rpm",
                            "terrain": "Plano, sem subidas"
                        }
                    }
                ]
            },
            "Saturday": {
                "day": "Sábado",
                "workouts": [
                    {
                        "type": "Musculação",
                        "focus": "Potência + Funcional",
                        "duration": 60,
                        "exercises": [
                            "Kettlebell Swing: 5x20",
                            "Box Jump: 4x8",
                            "Burpees: 4x12",
                            "Battle Rope: 4x30s",
                            "Wall Ball: 3x15"
                        ]
                    },
                    {
                        "type": "MTB Longo",
                        "focus": "Endurance Longo Z1-Z2",
                        "duration": 180 + int(progress * 60),  # 3-4h progressivo
                        "details": {
                            "warmup": "30min Z1",
                            "main": f"{150 + int(progress * 60)}min Z2 (116-133 bpm)",
                            "cooldown": "30min Z1",
                            "cadence": "80-90 rpm",
                            "technique": "Single track, curvas, subidas técnicas"
                        }
                    }
                ]
            },
            "Sunday": {
                "day": "Domingo",
                "workouts": [
                    {
                        "type": "Corrida Longa",
                        "focus": "Longão Z1-Z2",
                        "duration": 60 + int(progress * 30),  # 1h-1h30
                        "details": {
                            "warmup": "15min Z1",
                            "main": f"{30 + int(progress * 30)}min Z2",
                            "cooldown": "15min Z1",
                            "rpe": "5-6/10"
                        }
                    },
                    {
                        "type": "Recuperação",
                        "focus": "Flexibilidade + Liberação",
                        "duration": 30,
                        "details": {
                            "foam_rolling": "15min corpo todo",
                            "stretching": "15min principais grupos"
                        }
                    }
                ]
            }
        }
        
        return weekly_plan

# Interface Principal
def main():
    load_css()
    
    # Inicialização
    db = DatabaseManager()
    athlete = AthleteProfile()
    
    # Sidebar Premium
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 2rem 0;">
            <h2 style="color: #667eea; margin: 0;">🚴‍♂️ PRO AI COACH</h2>
            <p style="color: #a0a0b0; font-size: 0.9rem;">Sistema de Elite v2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Perfil do Atleta
        st.markdown("---")
        st.markdown("### 👤 Perfil do Atleta")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Idade", "29")
            st.metric("Altura", "1.90m")
            st.metric("FC Máx", "178 bpm")
        with col2:
            st.metric("Peso", "110kg")
            st.metric("Meta", "78kg")
            st.metric("Prazo", "31/12/26")
        
        # Progresso
        st.markdown("---")
        st.markdown("### 📊 Progresso")
        
        months_to_goal = 30  # Jul/2026 até Dez/2026
        weight_lost = 0  # Inicial
        progress_pct = 0
        
        st.markdown(f"""
        <div class="custom-progress">
            <p style="color: #a0a0b0; margin-bottom: 0.5rem;">Rumo aos 78kg</p>
            <div class="progress-fill" style="width: {progress_pct}%;"></div>
            <p style="color: #a0a0b0; margin-top: 0.5rem;">{weight_lost:.1f}kg / 32kg</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Navegação
        st.markdown("---")
        st.markdown("### 🎯 Navegação")
        
        page = st.radio(
            "",
            ["🏠 Dashboard", "🍽️ Plano Alimentar", "🏋️ Plano de Treinos", 
             "❤️ Zonas Cardíacas", "📈 Progressão 6 Meses", "📝 Registros Diários",
             "📊 Análises", "📅 Calendário de Provas"],
            label_visibility="collapsed"
        )
    
    # Conteúdo Principal
    if "Dashboard" in page:
        render_dashboard(athlete, db)
    elif "Alimentar" in page:
        render_nutrition_plan(athlete, db)
    elif "Treinos" in page:
        render_workout_plan(athlete, db)
    elif "Cardíacas" in page:
        render_heart_zones(athlete)
    elif "Progressão" in page:
        render_progression(athlete)
    elif "Registros" in page:
        render_daily_logs(db)
    elif "Análises" in page:
        render_analytics(db)
    elif "Calendário" in page:
        render_race_calendar()

def render_dashboard(athlete, db):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">🏆 Dashboard do Atleta</h1>
        <p class="hero-subtitle">Acompanhe sua evolução em tempo real</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Métricas Principais
    col1, col2, col3, col4 = st.columns(4)
    
    metrics = [
        ("🎯 Peso Atual", "110 kg", "Meta: 78 kg", "↓ 32 kg"),
        ("🔥 Calorias Hoje", "2.450", "Déficit: 500", "Meta: 2.580"),
        ("⏱️ Treino Hoje", "1h30", "Z2 Endurance", "RPE: 6/10"),
        ("💪 Adesão", "94%", "+2% vs semana", "Excelente")
    ]
    
    for col, (label, value, sublabel, delta) in zip([col1, col2, col3, col4], metrics):
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
                <div style="color: #a0a0b0;">{sublabel}</div>
                <div class="metric-delta">{delta}</div>
            </div>
            """, unsafe_allow_html=True)
    
    # Gráficos
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📈 Projeção de Peso")
        
        months = ['Jul/26', 'Ago/26', 'Set/26', 'Out/26', 'Nov/26', 'Dez/26']
        projected = [110, 107, 104, 101, 98, 95]  # Projeção realista até Dez/2026
        target = [110, 106, 102, 98, 94, 90]  # Meta ajustada
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=months, y=projected, mode='lines+markers',
                                name='Projeção Realista', line=dict(color='#667eea', width=3)))
        fig.add_trace(go.Scatter(x=months, y=target, mode='lines+markers',
                                name='Meta Intermediária', line=dict(color='#764ba2', width=2, dash='dash')))
        fig.add_hline(y=78, line_dash="dot", line_color="green", 
                     annotation_text="Meta Final: 78kg")
        
        fig.update_layout(
            template='plotly_dark',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            hovermode='x unified'
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("### 🏋️ Distribuição Semanal")
        
        workout_types = ['Musculação', 'Ciclismo', 'Corrida', 'Flexibilidade']
        hours = [6, 6.5, 3, 0.5]
        colors = ['#FF4B4B', '#0068C9', '#00C853', '#FFD700']
        
        fig = go.Figure(data=[go.Pie(labels=workout_types, values=hours, 
                                     hole=.4, marker_colors=colors)])
        fig.update_layout(
            template='plotly_dark',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Resumo Semanal
    st.markdown("### 📅 Resumo da Semana")
    
    days = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']
    durations = [150, 105, 150, 105, 120, 240, 105]  # Minutos
    
    fig = go.Figure(data=[
        go.Bar(name='Ciclismo', x=days, y=[90, 0, 90, 0, 60, 210, 0], 
               marker_color='#0068C9'),
        go.Bar(name='Musculação', x=days, y=[60, 60, 60, 60, 60, 60, 0], 
               marker_color='#FF4B4B'),
        go.Bar(name='Corrida', x=days, y=[0, 45, 0, 45, 0, 0, 75], 
               marker_color='#00C853')
    ])
    
    fig.update_layout(
        barmode='stack',
        template='plotly_dark',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        yaxis_title='Minutos'
    )
    st.plotly_chart(fig, use_container_width=True)

def render_nutrition_plan(athlete, db):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">🍽️ Plano Nutricional Premium</h1>
        <p class="hero-subtitle">Nutrição de precisão para alta performance</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Cálculos Nutricionais
    calc = NutritionCalculator()
    macros = calc.calculate_macros(110, 78, 30)
    
    # Macros Overview
    st.markdown("### 🎯 Metas Diárias Personalizadas")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🎯 Calorias Alvo", f"{macros['target_calories']} kcal", 
                 f"-{macros['deficit']} kcal")
    with col2:
        st.metric("💪 Proteínas", f"{macros['protein_g']}g", 
                 f"{macros['protein_pct']}%")
    with col3:
        st.metric("🍚 Carboidratos", f"{macros['carbs_g']}g", 
                 f"{macros['carbs_pct']}%")
    with col4:
        st.metric("🥑 Gorduras", f"{macros['fat_g']}g", 
                 f"{macros['fat_pct']}%")
    
    # Plano de Refeições
    st.markdown("---")
    st.markdown("### 📋 Refeições do Dia")
    
    meals_plan = {
        "Café da Manhã (06:30)": {
            "items": [
                ("Aveia em Flocos", 40),
                ("Whey Protein", 30),
                ("Banana", 120),
                ("Leite Desnatado", 200),
                ("Semente de Chia", 10)
            ],
            "timing": "Pré-treino",
            "notes": "Consumir 1h antes do treino"
        },
        "Lanche Pré-Treino (09:00)": {
            "items": [
                ("Pão Integral", 50),
                ("Ovo Inteiro", 100),
                ("Suco de Laranja Natural", 100)
            ],
            "timing": "30min antes",
            "notes": "Refeição leve para não pesar"
        },
        "Almoço (12:30)": {
            "items": [
                ("Peito de Frango Grelhado", 200),
                ("Arroz Integral Cozido", 150),
                ("Brócolis Cozido", 200),
                ("Azeite de Oliva", 15)
            ],
            "timing": "Pós-treino",
            "notes": "Maior refeição do dia"
        },
        "Lanche da Tarde (16:00)": {
            "items": [
                ("Iogurte Grego Natural", 200),
                ("Aveia em Flocos", 30),
                ("Maçã", 150),
                ("Castanha do Pará", 20)
            ],
            "timing": "Recuperação",
            "notes": "Energia para segundo treino"
        },
        "Jantar (19:30)": {
            "items": [
                ("Salmão", 200),
                ("Batata Doce Cozida", 150),
                ("Espinafre", 100),
                ("Tomate", 100)
            ],
            "timing": "Recuperação noturna",
            "notes": "Foco em proteínas e antioxidantes"
        },
        "Ceia (22:00)": {
            "items": [
                ("Queijo Cottage", 100),
                ("Amêndoas", 15)
            ],
            "timing": "Anti-catabólico",
            "notes": "Caseína para liberação lenta"
        }
    }
    
    for meal_name, meal_data in meals_plan.items():
        with st.expander(f"**{meal_name}** - {meal_data['timing']}", expanded=False):
            st.markdown(f"*{meal_data['notes']}*")
            
            total_cal = 0
            total_prot = 0
            
            for food, qty in meal_data['items']:
                # Buscar dados do banco
                c = db.conn.cursor()
                c.execute("SELECT * FROM food_database WHERE name=?", (food,))
                result = c.fetchone()
                
                if result:
                    cal = (result[2] * qty / 100)
                    prot = (result[3] * qty / 100)
                    total_cal += cal
                    total_prot += prot
                    
                    st.markdown(f"""
                    <div style="display: flex; justify-content: space-between; 
                              padding: 0.5rem; background: rgba(30,30,47,0.5); 
                              border-radius: 8px; margin: 0.3rem 0;">
                        <span>🍽️ {food}</span>
                        <span>{qty}g</span>
                        <span>{cal:.0f} kcal</span>
                        <span>{prot:.1f}g prot</span>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div style="text-align: right; color: #667eea; font-weight: bold;">
                Total: {total_cal:.0f} kcal | {total_prot:.0f}g proteína
            </div>
            """, unsafe_allow_html=True)
    
    # Suplementação
    st.markdown("---")
    st.markdown("### 💊 Suplementação Recomendada")
    
    supplements = [
        ("Whey Protein", "30g pós-treino", "Recuperação muscular rápida"),
        ("Creatina", "5g/dia", "Força e potência muscular"),
        ("Ômega-3", "2g/dia", "Anti-inflamatório, cardiovascular"),
        ("Vitamina D", "2000 UI/dia", "Imunidade e saúde óssea"),
        ("Magnésio", "400mg antes dormir", "Recuperação e sono"),
        ("Multivitamínico", "1 dose manhã", "Micronutrientes essenciais"),
        ("BCAA", "5g intra-treino", "Preservação muscular em treinos longos"),
        ("Maltodextrina", "30-60g/h (treinos >2h)", "Energia sustentada")
    ]
    
    for sup, dose, benefit in supplements:
        st.markdown(f"""
        <div class="food-card">
            <strong>{sup}</strong> - {dose}<br>
            <small style="color: #a0a0b0;">{benefit}</small>
        </div>
        """, unsafe_allow_html=True)

def render_workout_plan(athlete, db):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">🏋️ Programa de Treinamento</h1>
        <p class="hero-subtitle">Periodização científica para resultados máximos</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Seleção de semana
    week = st.slider("Semana do Programa", 1, 26, 1)
    
    planner = WorkoutPlanner()
    weekly_plan = planner.generate_weekly_plan(week)
    
    # Métricas da Semana
    total_duration = sum(
        workout['duration'] 
        for day in weekly_plan.values() 
        for workout in day['workouts']
    )
    total_hours = total_duration / 60
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("⏱️ Volume Total", f"{total_hours:.1f}h")
    with col2:
        st.metric("🏋️ Musculação", "6 sessões")
    with col3:
        st.metric("🚴 Ciclismo", "4 sessões")
    with col4:
        st.metric("🏃 Corrida", "3 sessões")
    
    st.markdown("---")
    
    # Plano Diário
    for day_key, day_data in weekly_plan.items():
        st.markdown(f"### {day_data['day']}")
        
        for workout in day_data['workouts']:
            workout_type = workout['type'].lower()
            
            if 'musculação' in workout_type:
                card_class = "musculacao"
                emoji = "💪"
            elif 'ciclismo' in workout_type or 'mtb' in workout_type:
                card_class = "ciclismo"
                emoji = "🚴"
            elif 'corrida' in workout_type:
                card_class = "corrida"
                emoji = "🏃"
            else:
                card_class = ""
                emoji = "🧘"
            
            st.markdown(f"""
            <div class="workout-card {card_class}">
                <h4>{emoji} {workout['focus']}</h4>
                <p><strong>Duração:</strong> {workout['duration']} minutos</p>
            </div>
            """, unsafe_allow_html=True)
            
            if 'exercises' in workout:
                with st.expander("Ver exercícios"):
                    for exercise in workout['exercises']:
                        st.write(f"• {exercise}")
            
            if 'details' in workout:
                with st.expander("Ver detalhes do treino"):
                    for key, value in workout['details'].items():
                        st.write(f"**{key.replace('_', ' ').title()}:** {value}")
        
        st.markdown("---")

def render_heart_zones(athlete):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">❤️ Zonas de Frequência Cardíaca</h1>
        <p class="hero-subtitle">Baseado na sua FC Máx de 178 bpm</p>
    </div>
    """, unsafe_allow_html=True)
    
    zones = {
        "Z1": {"name": "Recuperação Ativa", "range": "<116 bpm", "pct": "<65%", 
               "use": "Aquecimento, volta calma, dias de recuperação",
               "benefit": "Recuperação, técnica, base aeróbica leve",
               "rpe": "1-3/10"},
        "Z2": {"name": "Endurance Aeróbica", "range": "116-133 bpm", "pct": "66-75%", 
               "use": "Treinos longos, base, queima de gordura",
               "benefit": "Eficiência cardiovascular, oxidação de gordura",
               "rpe": "4-5/10"},
        "Z3": {"name": "Tempo/Ritmo", "range": "134-151 bpm", "pct": "76-85%", 
               "use": "Ritmo de prova, sustentação",
               "benefit": "Lactato threshold, endurance muscular",
               "rpe": "6-7/10"},
        "Z4": {"name": "Limiar Anaeróbico", "range": "152-163 bpm", "pct": "86-92%", 
               "use": "Intervalados, FTP, subidas fortes",
               "benefit": "Tolerância ao lactato, potência sustentada",
               "rpe": "7-8/10"},
        "Z5": {"name": "VO2 Máximo", "range": "164-178 bpm", "pct": "93-100%", 
               "use": "Sprints, ataques, tiros curtos",
               "benefit": "Potência máxima, capacidade anaeróbica",
               "rpe": "9-10/10"}
    }
    
    for zone_code, zone_data in zones.items():
        zone_class = f"zone-{zone_code.lower()}"
        
        st.markdown(f"""
        <div class="food-card">
            <span class="zone-badge {zone_class}">{zone_code}</span>
            <strong>{zone_data['name']}</strong><br>
            <div style="margin-top: 1rem;">
                <strong>Faixa:</strong> {zone_data['range']} ({zone_data['pct']})<br>
                <strong>RPE:</strong> {zone_data['rpe']}<br>
                <strong>Uso Principal:</strong> {zone_data['use']}<br>
                <strong>Benefício:</strong> {zone_data['benefit']}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # Gráfico Visual
    st.markdown("### 📊 Distribuição Visual")
    
    fig = go.Figure()
    
    zone_colors = ['#00C853', '#64DD17', '#FFD700', '#FF9100', '#FF1744']
    zone_names = ['Z1', 'Z2', 'Z3', 'Z4', 'Z5']
    
    for i, (name, color) in enumerate(zip(zone_names, zone_colors)):
        fig.add_trace(go.Bar(
            name=name,
            x=['Zonas'],
            y=[1],
            marker_color=color,
            text=name,
            textposition='inside'
        ))
    
    fig.update_layout(
        barmode='stack',
        showlegend=False,
        height=100,
        template='plotly_dark',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)'
    )
    st.plotly_chart(fig, use_container_width=True)

def render_progression(athlete):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">📈 Progressão de 6 Meses</h1>
        <p class="hero-subtitle">Periodização completa Julho-Dezembro 2026</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Linha do tempo
    months = [
        {
            "month": "Julho 2026",
            "phase": "Adaptação",
            "weeks": "1-4",
            "volume": "8-10h/semana",
            "intensity": "Z1-Z2: 70% | Z3-Z4: 30%",
            "focus": "Adaptação muscular, base aeróbica, técnica",
            "weight_goal": "110kg → 107kg (-3kg)",
            "nutrition": "Déficit 300-400 kcal, adaptação metabólica",
            "tests": "Teste FTP, avaliação física completa"
        },
        {
            "month": "Agosto 2026",
            "phase": "Base",
            "weeks": "5-8",
            "volume": "10-12h/semana",
            "intensity": "Z1-Z2: 65% | Z3-Z4: 35%",
            "focus": "Aumento volume, endurance, força na musculação",
            "weight_goal": "107kg → 104.5kg (-2.5kg)",
            "nutrition": "Déficit 400-500 kcal, periodização de carbs",
            "tests": "Reavaliação composição corporal"
        },
        {
            "month": "Setembro 2026",
            "phase": "Construção 1",
            "weeks": "9-12",
            "volume": "12-14h/semana",
            "intensity": "Z1-Z2: 60% | Z3-Z4: 30% | Z5: 10%",
            "focus": "Introdução VO2 Max, intervalados intensos",
            "weight_goal": "104.5kg → 102kg (-2.5kg)",
            "nutrition": "Déficit 500 kcal, nutrição peri-treino",
            "tests": "Teste de campo 40km MTB"
        },
        {
            "month": "Outubro 2026",
            "phase": "Construção 2",
            "weeks": "13-16",
            "volume": "14-15h/semana",
            "intensity": "Z1-Z2: 55% | Z3-Z4: 35% | Z5: 10%",
            "focus": "Pico volume, especificidade XCM, brick training",
            "weight_goal": "102kg → 99.5kg (-2.5kg)",
            "nutrition": "Déficit 500-600 kcal, máxima eficiência",
            "tests": "Simulado XCM 3h"
        },
        {
            "month": "Novembro 2026",
            "phase": "Específico",
            "weeks": "17-20",
            "volume": "12-14h/semana",
            "intensity": "Z1-Z2: 50% | Z3-Z4: 40% | Z5: 10%",
            "focus": "Redução volume, aumento intensidade, ritmo prova",
            "weight_goal": "99.5kg → 97.5kg (-2kg)",
            "nutrition": "Déficit 400-500 kcal, teste nutrição prova",
            "tests": "Simulado completo duatlo"
        },
        {
            "month": "Dezembro 2026",
            "phase": "Competição/Taper",
            "weeks": "21-24",
            "volume": "10-12h/semana",
            "intensity": "Z1-Z2: 60% | Z3-Z4: 25% | Z5: 15%",
            "focus": "Manutenção forma, frescor para provas",
            "weight_goal": "97.5kg → 95kg (-2.5kg)*",
            "nutrition": "Ajuste fino, supercompensação",
            "tests": "Provas alvo"
        }
    ]
    
    for month_data in months:
        with st.expander(f"**{month_data['month']}** - Fase {month_data['phase']} | Meta: {month_data['weight_goal']}", 
                        expanded=month_data['month'] == "Julho 2026"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"""
                **Semanas:** {month_data['weeks']}  
                **Volume:** {month_data['volume']}  
                **Intensidade:** {month_data['intensity']}  
                **Foco:** {month_data['focus']}
                """)
            
            with col2:
                st.markdown(f"""
                **Nutrição:** {month_data['nutrition']}  
                **Testes:** {month_data['tests']}
                """)
    
    st.info("""
    *A meta de 78kg em 31/12/2026 é alcançada com um déficit calórico consistente. 
    O ritmo de perda de peso é progressivo e sustentável, preservando massa muscular e performance.
    """)
    
    # Gráfico de progressão de carga
    st.markdown("### 📊 Progressão de Carga de Treino")
    
    weeks = list(range(1, 25))
    volume = [9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14, 14.5, 14.5, 15, 14.5, 
              14, 14, 13, 13, 12, 12, 11, 11]
    intensity = [30, 30, 35, 35, 40, 40, 40, 40, 45, 45, 45, 45, 50, 50, 50, 50, 
                55, 55, 55, 55, 50, 50, 45, 45]  # % Z3-Z5
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Scatter(x=weeks, y=volume, name="Volume (h/sem)", line=dict(color="#667eea", width=3)),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Scatter(x=weeks, y=intensity, name="Intensidade (%Z3-Z5)", 
                  line=dict(color="#FF4B4B", width=3)),
        secondary_y=True
    )
    
    fig.update_layout(
        template='plotly_dark',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        hovermode='x unified'
    )
    
    fig.update_xaxes(title_text="Semanas")
    fig.update_yaxes(title_text="Horas/Semana", secondary_y=False)
    fig.update_yaxes(title_text="% Intensidade", secondary_y=True)
    
    st.plotly_chart(fig, use_container_width=True)

def render_daily_logs(db):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">📝 Registros Diários</h1>
        <p class="hero-subtitle">Monitore cada detalhe da sua jornada</p>
    </div>
    """, unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["💪 Treino", "🍽️ Alimentação", "⚖️ Peso"])
    
    with tab1:
        st.subheader("Registrar Treino")
        
        col1, col2 = st.columns(2)
        
        with col1:
            workout_date = st.date_input("Data", datetime.now(), key="workout_date")
            workout_type = st.selectbox("Tipo", ["Musculação", "MTB", "Corrida", "Brick", "Flexibilidade"])
            duration = st.number_input("Duração (min)", 0, 480, 60)
            distance = st.number_input("Distância (km)", 0.0, 200.0, 0.0, 0.1)
        
        with col2:
            avg_hr = st.number_input("FC Média", 0, 220, 135)
            max_hr = st.number_input("FC Máxima", 0, 220, 155)
            cadence = st.number_input("Cadência Média", 0, 120, 85)
            zone = st.selectbox("Zona Principal", ["Z1", "Z2", "Z3", "Z4", "Z5"])
            rpe = st.slider("RPE (Escala de Borg)", 1, 10, 6)
        
        feeling = st.select_slider("Sensação", ["😫 Muito Ruim", "😔 Ruim", "😐 Neutro", 
                                                "🙂 Bom", "😀 Muito Bom", "🤩 Excelente"])
        
        notes = st.text_area("Observações", "Treino conforme planejado")
        
        if st.button("💾 Salvar Treino", use_container_width=True):
            # Estimativa de calorias
            calorie_factors = {
                "Musculação": 6,
                "MTB": 11,
                "Corrida": 12,
                "Brick": 11.5,
                "Flexibilidade": 3
            }
            calories = duration * calorie_factors.get(workout_type, 8)
            
            c = db.conn.cursor()
            c.execute('''
                INSERT INTO workouts (date, type, duration_min, distance_km, 
                avg_hr, max_hr, avg_cadence, zone, rpe, calories_burned, feeling, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (workout_date, workout_type, duration, distance, avg_hr, max_hr, 
                 cadence, zone, rpe, calories, feeling, notes))
            db.conn.commit()
            
            st.success(f"✅ Treino salvo! ~{calories:.0f} calorias queimadas")
    
    with tab2:
        st.subheader("Registrar Refeição")
        
        col1, col2 = st.columns(2)
        
        with col1:
            meal_date = st.date_input("Data", datetime.now(), key="meal_date")
            meal_type = st.selectbox("Refeição", [
                "Café da Manhã", "Lanche Manhã", "Almoço", 
                "Lanche Tarde", "Jantar", "Ceia", "Intra-treino"
            ])
            
            # Buscar alimentos do banco
            c = db.conn.cursor()
            c.execute("SELECT name FROM food_database ORDER BY name")
            foods = [row[0] for row in c.fetchall()]
            
            food_name = st.selectbox("Alimento", foods)
            portion = st.number_input("Porção (g)", 10, 1000, 100, 10)
        
        with col2:
            # Buscar dados do alimento
            c.execute("SELECT * FROM food_database WHERE name=?", (food_name,))
            food_data = c.fetchone()
            
            if food_data:
                cal_portion = (food_data[2] * portion / 100)
                prot_portion = (food_data[3] * portion / 100)
                carbs_portion = (food_data[4] * portion / 100)
                fat_portion = (food_data[5] * portion / 100)
                fiber_portion = (food_data[6] * portion / 100) if food_data[6] else 0
                
                st.info(f"**{food_name}** ({portion}g)")
                st.metric("Calorias", f"{cal_portion:.0f} kcal")
                col1_1, col2_2 = st.columns(2)
                col1_1.metric("Proteínas", f"{prot_portion:.1f}g")
                col1_1.metric("Carboidratos", f"{carbs_portion:.1f}g")
                col2_2.metric("Gorduras", f"{fat_portion:.1f}g")
                col2_2.metric("Fibras", f"{fiber_portion:.1f}g")
        
        if st.button("💾 Registrar Refeição", use_container_width=True):
            c = db.conn.cursor()
            c.execute('''
                INSERT INTO meals (date, meal_type, food_name, portion_g, 
                calories, protein_g, carbs_g, fat_g, fiber_g)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (meal_date, meal_type, food_name, portion, cal_portion, 
                 prot_portion, carbs_portion, fat_portion, fiber_portion))
            db.conn.commit()
            
            st.success("✅ Refeição registrada!")
    
    with tab3:
        st.subheader("Registrar Peso")
        
        col1, col2 = st.columns(2)
        
        with col1:
            weight_date = st.date_input("Data", datetime.now(), key="weight_date")
            weight = st.number_input("Peso (kg)", 70.0, 120.0, 110.0, 0.1)
        
        with col2:
            body_fat = st.number_input("% Gordura Corporal (opcional)", 0.0, 50.0, 25.0, 0.1)
            weight_notes = st.text_area("Notas", "")
        
        if st.button("⚖️ Registrar Peso", use_container_width=True):
            c = db.conn.cursor()
            c.execute('''
                INSERT INTO weight_log (date, weight_kg, body_fat_pct, notes)
                VALUES (?, ?, ?, ?)
            ''', (weight_date, weight, body_fat, weight_notes))
            db.conn.commit()
            
            st.success("✅ Peso registrado!")

def render_analytics(db):
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">📊 Análises Avançadas</h1>
        <p class="hero-subtitle">Insights baseados em dados</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Dados de exemplo para demonstração
    dates = pd.date_range('2026-07-01', periods=180, freq='D')
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📈 Tendência de Peso")
        
        # Simulação de dados
        np.random.seed(42)
        weights = 110 - np.linspace(0, 15, 180) + np.random.normal(0, 0.5, 180)
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=dates, y=weights, mode='lines',
                                name='Peso Real', line=dict(color='#667eea', width=2)))
        fig.add_hline(y=78, line_dash="dash", line_color="#00C853",
                     annotation_text="Meta: 78kg")
        
        fig.update_layout(
            template='plotly_dark',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            yaxis_title='Peso (kg)'
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("### 🔥 Calorias Semanais")
        
        weeks = list(range(1, 27))
        calories_burned = np.random.normal(5000, 500, 26) + np.linspace(0, 2000, 26)
        calories_consumed = calories_burned - np.random.normal(3500, 300, 26)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=weeks, y=calories_burned, name='Gasto',
                            marker_color='#FF4B4B'))
        fig.add_trace(go.Bar(x=weeks, y=calories_consumed, name='Consumo',
                            marker_color='#667eea'))
        
        fig.update_layout(
            template='plotly_dark',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            barmode='group'
        )
        st.plotly_chart(fig, use_container_width=True)

def render_race_calendar():
    st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">📅 Calendário de Provas</h1>
        <p class="hero-subtitle">Preparação para competições alvo</p>
    </div>
    """, unsafe_allow_html=True)
    
    races = [
        {
            "date": "Outubro 2026",
            "event": "Desafio de MTB",
            "type": "XCM 60km",
            "goal": "Completar em < 3h30",
            "priority": "Média"
        },
        {
            "date": "Novembro 2026",
            "event": "Duatlo Regional",
            "type": "5km/30km/5km",
            "goal": "Top 50% categoria",
            "priority": "Alta"
        },
        {
            "date": "Dezembro 2026",
            "event": "Maratona MTB Principal",
            "type": "XCM 80km",
            "goal": "Completar em < 4h30",
            "priority": "Máxima"
        }
    ]
    
    for race in races:
        priority_color = {
            "Máxima": "#FF1744",
            "Alta": "#FF9100",
            "Média": "#FFD700"
        }.get(race['priority'], '#00C853')
        
        st.markdown(f"""
        <div class="food-card" style="border-left: 5px solid {priority_color};">
            <div style="display: flex; justify-content: space-between;">
                <h3>{race['event']}</h3>
                <span class="zone-badge" style="background-color: {priority_color}; 
                      color: white;">{race['priority']}</span>
            </div>
            <p><strong>Data:</strong> {race['date']}</p>
            <p><strong>Tipo:</strong> {race['type']}</p>
            <p><strong>Meta:</strong> {race['goal']}</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()