"""
SUPPLIER PERFORMANCE ANALYSIS — UI COMPONENT FRAMEWORK
Module: src/ui.py
Styling: Tailwind CSS CDN injected via streamlit.components.v1.html
Theme: Enterprise Slate & Navy Blue with Emerald, Amber, and Rose Accents
"""

import streamlit as st
import streamlit.components.v1 as components

def render_tailwind_header(title: str, subtitle: str, badge_text: str = "PRODUCTION READY"):
    """
    Renders an enterprise gradient header card using Tailwind CSS in an iframe.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <script src="https://cdn.tailwindcss.com"></script>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }}
        </style>
    </head>
    <body class="bg-transparent m-0 p-1">
        <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-900 text-white rounded-xl shadow-lg p-6 border border-slate-700/50 flex flex-col md:flex-row justify-between items-start md:items-center">
            <div>
                <div class="flex items-center space-x-3 mb-2">
                    <span class="bg-emerald-500/20 text-emerald-400 text-xs font-semibold px-2.5 py-0.5 rounded-full border border-emerald-500/30 uppercase tracking-wider">
                        {badge_text}
                    </span>
                    <span class="text-slate-400 text-xs">ABES MCA Capstone Project</span>
                </div>
                <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight text-white mb-1">
                    {title}
                </h1>
                <p class="text-slate-300 text-sm max-w-2xl">
                    {subtitle}
                </p>
            </div>
            <div class="mt-4 md:mt-0 flex items-center space-x-2 text-xs text-slate-300 bg-slate-800/60 px-3 py-2 rounded-lg border border-slate-700">
                <span class="h-2 w-2 rounded-full bg-emerald-400 animate-pulse"></span>
                <span>SQLite DB: Connected</span>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=135)

def render_kpi_cards(kpi_data: list):
    """
    Renders a responsive grid of modern executive KPI metric cards.
    kpi_data format: list of dicts with keys:
    [
        {
            "label": "Overall OTIF Rate",
            "value": "89.4%",
            "benchmark": "Target >= 95.0%",
            "status": "warning",  # "success", "warning", "danger", "neutral"
            "subtext": "Fulfillment punctuality"
        }, ...
    ]
    """
    cards_html = ""
    for item in kpi_data:
        status = item.get("status", "neutral")
        if status == "success":
            badge_color = "bg-emerald-50 text-emerald-700 border-emerald-200"
            border_accent = "border-t-4 border-t-emerald-500"
            dot_color = "bg-emerald-500"
        elif status == "warning":
            badge_color = "bg-amber-50 text-amber-700 border-amber-200"
            border_accent = "border-t-4 border-t-amber-500"
            dot_color = "bg-amber-500"
        elif status == "danger":
            badge_color = "bg-rose-50 text-rose-700 border-rose-200"
            border_accent = "border-t-4 border-t-rose-500"
            dot_color = "bg-rose-500"
        else:
            badge_color = "bg-slate-50 text-slate-700 border-slate-200"
            border_accent = "border-t-4 border-t-slate-500"
            dot_color = "bg-slate-400"

        cards_html += f"""
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-col justify-between {border_accent} hover:shadow-md transition-shadow">
            <div>
                <div class="flex justify-between items-center mb-2">
                    <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">{item['label']}</span>
                    <span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium border {badge_color}">
                        <span class="w-1.5 h-1.5 mr-1.5 rounded-full {dot_color}"></span>
                        {item.get('benchmark', '')}
                    </span>
                </div>
                <div class="text-3xl font-black text-slate-900 tracking-tight my-1">
                    {item['value']}
                </div>
            </div>
            <div class="text-xs text-slate-400 mt-3 pt-2 border-t border-slate-100 flex items-center justify-between">
                <span>{item.get('subtext', '')}</span>
            </div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <script src="https://cdn.tailwindcss.com"></script>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }}
        </style>
    </head>
    <body class="bg-transparent m-0 p-1">
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {cards_html}
        </div>
    </body>
    </html>
    """
    components.html(full_html, height=140)

def render_insight_box(title: str, message: str, alert_type: str = "info"):
    """
    Renders a styled analytical insight callout box.
    """
    if alert_type == "danger":
        theme = "bg-rose-50 border-rose-200 text-rose-900 icon-rose-500"
        icon = "⚠️"
    elif alert_type == "warning":
        theme = "bg-amber-50 border-amber-200 text-amber-900 icon-amber-500"
        icon = "⚡"
    elif alert_type == "success":
        theme = "bg-emerald-50 border-emerald-200 text-emerald-900 icon-emerald-500"
        icon = "✓"
    else:
        theme = "bg-blue-50 border-blue-200 text-blue-900 icon-blue-500"
        icon = "ℹ️"

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-1">
        <div class="rounded-lg border p-4 shadow-sm {theme}">
            <div class="flex items-start space-x-3">
                <span class="text-xl leading-none mt-0.5">{icon}</span>
                <div>
                    <h4 class="text-sm font-bold tracking-tight mb-1">{title}</h4>
                    <p class="text-xs leading-relaxed text-slate-700">{message}</p>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=95)
