# -*- coding: utf-8 -*-
"""
generate_diagrams.py
Generates professional, high-resolution (300 DPI) architecture and data-flow diagrams
matching the project's color palette (Navy #0D1B2A, Blue #1E3A5F, Gold #C9A84C, Slate #475569).
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches

fig_dir = r"d:\AVLANCE\Advocate\figures"
os.makedirs(fig_dir, exist_ok=True)

# Color Palette
NAVY = "#0D1B2A"
BLUE = "#1E3A5F"
LIGHT_BLUE = "#EBF3FB"
GOLD = "#C9A84C"
LIGHT_GOLD = "#FDF8ED"
SLATE = "#334155"
LIGHT_BG = "#F8FAFC"
BORDER_GRAY = "#CBD5E1"
WHITE = "#FFFFFF"

def draw_box(ax, x, y, w, h, text, title="", bg_color=WHITE, border_color=BORDER_GRAY, title_color=NAVY, text_color=SLATE, lw=1.5):
    # Shadow
    shadow = patches.FancyBboxPatch((x+0.05, y-0.05), w, h, boxstyle="round,pad=0.08,rounding_size=0.15",
                                    facecolor="#E2E8F0", edgecolor="none", zorder=1)
    ax.add_patch(shadow)
    # Main Box
    box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08,rounding_size=0.15",
                                facecolor=bg_color, edgecolor=border_color, linewidth=lw, zorder=2)
    ax.add_patch(box)
    if title:
        ax.text(x + w/2, y + h - 0.28, title, ha='center', va='center', fontsize=9.5, fontweight='bold', color=title_color, zorder=3)
    if text:
        ax.text(x + w/2, y + (h/2 if not title else (h/2 - 0.15)), text, ha='center', va='center', fontsize=8.2, color=text_color, zorder=3, multialignment='center')

def draw_arrow(ax, x1, y1, x2, y2, label="", color=GOLD, rad=0.0):
    connectionstyle = f"arc3,rad={rad}" if rad != 0.0 else "arc3"
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle="-|>", color=color, lw=1.8, mutation_scale=14, connectionstyle=connectionstyle), zorder=4)
    if label:
        mx = (x1 + x2) / 2
        my = (y1 + y2) / 2 + 0.12
        ax.text(mx, my, label, ha='center', va='center', fontsize=7.5, fontweight='bold', color=BLUE,
                bbox=dict(boxstyle="round,pad=0.2", facecolor=WHITE, edgecolor=BORDER_GRAY, lw=0.8), zorder=5)

# ==============================================================================
# Diagram 1: Figure 5.1 - Static JAMstack 3-Tier Architecture
# ==============================================================================
def create_fig_5_1():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    # Header Title
    ax.text(5, 4.7, "Figure 5.1: 3-Tier Static JAMstack Architecture & Zero-Server Model",
            ha='center', va='center', fontsize=12, fontweight='bold', color=NAVY)

    # Tier 1: Client Browser Tier
    draw_box(ax, 0.4, 0.6, 2.7, 3.6,
             "• Semantic HTML5 Markup\n• CSS3 Tokens & Keyframes\n• Vanilla ES6+ Interactive Controllers\n• SVG Interactive Scales Engine\n• LocalStorage State Cache",
             title="Tier 1: Litigant Client Tier\n(Desktop & Mobile Browsers)", bg_color=LIGHT_BLUE, border_color=BLUE, title_color=BLUE)

    # Tier 2: Edge CDN & Delivery Tier
    draw_box(ax, 3.65, 0.6, 2.7, 3.6,
             "• HTTPS / TLS 1.3 Encryption\n• Global Anycast Edge Network\n• Sub-Second Asset Caching\n• Zero-Server Attack Surface\n• 100% Uptime SLA Architecture",
             title="Tier 2: Edge CDN Tier\n(Cloudflare / GitHub Pages)", bg_color=LIGHT_GOLD, border_color=GOLD, title_color=GOLD)

    # Tier 3: Communication & API Tier
    draw_box(ax, 6.9, 0.6, 2.7, 3.6,
             "• Google Fonts CDN (Typography)\n• Google Maps Embed (Location)\n• WhatsApp Direct API (Instant Chat)\n• SMTP / Mailto Protocol (Briefs)\n• Zero Database Maintenance Cost",
             title="Tier 3: Service Gateway\n(Decoupled External APIs)", bg_color=LIGHT_BG, border_color=NAVY, title_color=NAVY)

    # Connecting Arrows
    draw_arrow(ax, 3.1, 2.7, 3.65, 2.7, label="HTTPS GET")
    draw_arrow(ax, 3.65, 1.7, 3.1, 1.7, label="Static Bundle")
    draw_arrow(ax, 6.35, 2.7, 6.9, 2.7, label="API Request")
    draw_arrow(ax, 6.9, 1.7, 6.35, 1.7, label="Font / Map Assets")

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "fig_5_1_architecture.png")
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Saved:", out_path)

# ==============================================================================
# Diagram 2: Figure 5.2 - Zero-Database Client Inquiry Flow
# ==============================================================================
def create_fig_5_2():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    ax.text(5, 4.7, "Figure 5.2: Zero-Database Client Consultation Inquiry & Docket Data Flow",
            ha='center', va='center', fontsize=12, fontweight='bold', color=NAVY)

    # Step 1: User Input
    draw_box(ax, 0.3, 1.2, 2.0, 2.7,
             "• Full Name & Phone\n• Legal Practice Domain\n• Incident / Dispute Date\n• Statement of Fact Brief",
             title="Step 1: Litigant Input\n(Petition Form #contact)", bg_color=WHITE, border_color=BORDER_GRAY)

    # Step 2: Client-side Validation
    draw_box(ax, 2.7, 1.2, 2.0, 2.7,
             "• Regex Phone Validation\n• Mandatory Field Scrutiny\n• HTML Sanitization\n• Modal Feedback Alert",
             title="Step 2: Script Engine\n(Form Validation)", bg_color=LIGHT_BLUE, border_color=BLUE, title_color=BLUE)

    # Step 3: Docket Generator
    draw_box(ax, 5.1, 1.2, 2.0, 2.7,
             "• Timestamp Hash\n• Prefix: BK-2026-\n• 5-Digit Unique Token\n• Jurat Registration Pill",
             title="Step 3: Docket Engine\n(Unique Ref Generation)", bg_color=LIGHT_GOLD, border_color=GOLD, title_color=GOLD)

    # Step 4: Direct Dispatch
    draw_box(ax, 7.5, 2.5, 2.2, 1.6,
             "Pre-populated docket brief\nsent directly to Advocate phone",
             title="WhatsApp Gateway\n(api.whatsapp.com)", bg_color="#E8F8F0", border_color="#10B981", title_color="#047857")

    draw_box(ax, 7.5, 0.5, 2.2, 1.6,
             "Structured legal brief\nsent to chamber email",
             title="Chamber Mail Intake\n(mailto: protocol)", bg_color="#EFF6FF", border_color="#3B82F6", title_color="#1D4ED8")

    # Arrows
    draw_arrow(ax, 2.3, 2.55, 2.7, 2.55, label="Submit")
    draw_arrow(ax, 4.7, 2.55, 5.1, 2.55, label="Verified")
    draw_arrow(ax, 7.1, 2.8, 7.5, 3.3, label="Mobile")
    draw_arrow(ax, 7.1, 2.3, 7.5, 1.3, label="Email")

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "fig_5_2_data_flow.png")
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Saved:", out_path)

# ==============================================================================
# Diagram 3: Figure 7.1 - Interactive Scales of Justice State Machine
# ==============================================================================
def create_fig_7_1():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    ax.text(5, 4.7, "Figure 7.1: Interactive Scales of Justice Beam Tipping State Machine",
            ha='center', va='center', fontsize=12, fontweight='bold', color=NAVY)

    # State 1: Equilibrium
    draw_box(ax, 0.4, 1.2, 2.6, 2.8,
             "• SVG Beam Angle = 0.0°\n• Left & Right Pans Equal\n• Status: Judicial Balance\n• Gold Glow = Idle Opacity 0.4",
             title="State 0: Neutral Balance\n(Resting Equilibrium)", bg_color=WHITE, border_color=BORDER_GRAY)

    # State 2: Mouse Enter / Tilt Trigger
    draw_box(ax, 3.7, 1.2, 2.6, 2.8,
             "• Litigant hovers practice folder\n• Reads data-tilt-weight (+/- 12°)\n• CSS cubic-bezier transition\n• Beam tips toward relevant scale",
             title="State 1: Dynamic Tilt\n(Practice Area Hover)", bg_color=LIGHT_GOLD, border_color=GOLD, title_color=GOLD)

    # State 3: Restore / Settling
    draw_box(ax, 7.0, 1.2, 2.6, 2.8,
             "• Cursor leaves folder rect\n• LERP ease-out returning to 0°\n• Micro-oscillation damping\n• Preserves GPU 60 FPS frame rate",
             title="State 2: Re-Equilibrium\n(Smooth Damped Return)", bg_color=LIGHT_BLUE, border_color=BLUE, title_color=BLUE)

    draw_arrow(ax, 3.0, 2.6, 3.7, 2.6, label="onmouseenter")
    draw_arrow(ax, 6.3, 2.6, 7.0, 2.6, label="onmouseleave")
    draw_arrow(ax, 7.0, 1.6, 3.0, 1.6, label="Reset Loop", rad=0.25)

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "fig_7_1_scales_state.png")
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Saved:", out_path)

# ==============================================================================
# Diagram 4: Figure 7.2 - Dynamic Bilingual Translation Flow
# ==============================================================================
def create_fig_7_2():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    ax.text(5, 4.7, "Figure 7.2: Dynamic Bilingual (English / Tamil) Translation Engine",
            ha='center', va='center', fontsize=12, fontweight='bold', color=NAVY)

    draw_box(ax, 0.4, 1.2, 2.0, 2.7,
             "• Click 'Tamil (TA)' or 'English'\n• Event listener fires\n• setLanguage(lang) called\n• Active pill UI toggle",
             title="Step 1: User Event\n(Language Switcher)", bg_color=WHITE, border_color=BORDER_GRAY)

    draw_box(ax, 2.8, 1.2, 2.1, 2.7,
             "• translations['en'] / ['ta']\n• Key-value string dictionary\n• Cultural Tamil legal jurat\n• Zero external API request",
             title="Step 2: Dictionary\n(Bilingual Lexicon)", bg_color=LIGHT_GOLD, border_color=GOLD, title_color=GOLD)

    draw_box(ax, 5.3, 1.2, 2.1, 2.7,
             "• document.querySelectorAll\n• Filter [data-i18n] nodes\n• Preserves HTML structure\n• Instant textContent replace",
             title="Step 3: DOM Engine\n(Node Traversal)", bg_color=LIGHT_BLUE, border_color=BLUE, title_color=BLUE)

    draw_box(ax, 7.7, 1.2, 1.9, 2.7,
             "• Header & Hero strings\n• Practice Area dossiers\n• Notary affidavit text\n• Statutory disclaimers",
             title="Step 4: Output\n(Rendered UI)", bg_color="#FAF5FF", border_color="#A855F7", title_color="#7E22CE")

    draw_arrow(ax, 2.4, 2.55, 2.8, 2.55, label="Trigger")
    draw_arrow(ax, 4.9, 2.55, 5.3, 2.55, label="Lookup")
    draw_arrow(ax, 7.4, 2.55, 7.7, 2.55, label="Update")

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "fig_7_2_translation_flow.png")
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Saved:", out_path)

# ==============================================================================
# Diagram 5: Figure 9.1 - CI/CD & Zero-Cost Deployment Pipeline
# ==============================================================================
def create_fig_9_1():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    ax.text(5, 4.7, "Figure 9.1: Continuous Integration & Zero-Cost Web Hosting Pipeline",
            ha='center', va='center', fontsize=12, fontweight='bold', color=NAVY)

    draw_box(ax, 0.3, 1.2, 2.0, 2.7,
             "• Vanilla HTML/CSS/JS\n• Zero framework bloat\n• Staging sandbox audit\n• Git branch management",
             title="Phase 1: Local Code\n(Source Repository)", bg_color=WHITE, border_color=BORDER_GRAY)

    draw_box(ax, 2.7, 1.2, 2.0, 2.7,
             "• Automated HTML5 lint\n• Google Lighthouse check\n• Mobile responsive audit\n• BCI Rule 36 compliance",
             title="Phase 2: QA Audit\n(Quality Gates)", bg_color=LIGHT_BLUE, border_color=BLUE, title_color=BLUE)

    draw_box(ax, 5.1, 1.2, 2.0, 2.7,
             "• Push to main branch\n• Automated GitHub Action\n• Instant static artifact build\n• Zero compilation overhead",
             title="Phase 3: Deploy\n(GitHub Pages CI)", bg_color=LIGHT_GOLD, border_color=GOLD, title_color=GOLD)

    draw_box(ax, 7.5, 1.2, 2.2, 2.7,
             "• Global Edge CDN delivery\n• Automated SSL Certificate\n• 0 ms cold start latency\n• Rs. 0/mo maintenance cost",
             title="Phase 4: Live Chamber\n(Global Litigant Access)", bg_color="#E8F8F0", border_color="#10B981", title_color="#047857")

    draw_arrow(ax, 2.3, 2.55, 2.7, 2.55, label="Commit")
    draw_arrow(ax, 4.7, 2.55, 5.1, 2.55, label="Verified")
    draw_arrow(ax, 7.1, 2.55, 7.5, 2.55, label="Publish")

    plt.tight_layout()
    out_path = os.path.join(fig_dir, "fig_9_1_deployment.png")
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    plt.close()
    print("Saved:", out_path)

if __name__ == "__main__":
    create_fig_5_1()
    create_fig_5_2()
    create_fig_7_1()
    create_fig_7_2()
    create_fig_9_1()
    print("All architecture diagrams successfully generated!")
