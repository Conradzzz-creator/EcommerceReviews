import re
from collections import Counter
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Amazon Reviews Analytics", page_icon="📊", layout="wide")

# Application styling: dark canvas, light-blue headings, green controls, teal chart panels.
st.markdown("""
<style>
:root {
  --app-green: #228B22;
  --app-black: #000000;
  --app-blue: #9DDCFF;
  --app-grey: #C7C7C7;
}
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"],
[data-testid="stSidebar"], [data-testid="stSidebar"] > div,
section[data-testid="stSidebar"] {
  background-color: var(--app-black) !important;
  color: #FFFFFF;
}
/* Keep the sidebar's widgets and text legible on the black background. */
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
  color: #FFFFFF;
}
html, body, [class*="st-"], p, li, label, [data-testid="stMarkdownContainer"] {
  color: #FFFFFF;
}
h1, h2, h3, h4, h5, h6, [data-testid="stHeader"] h1 {
  color: var(--app-blue) !important;
}
[data-testid="stCaptionContainer"], .stCaption, [data-testid="stTooltipContent"] {
  color: var(--app-grey) !important;
}
/* Green radio selection, action buttons and multiselect chips. */
input[type="radio"], input[type="checkbox"] { accent-color: var(--app-green) !important; }
[data-testid="stRadio"] [role="radio"][aria-checked="true"] > div,
[data-testid="stRadio"] label[data-checked="true"] > div {
  background-color: var(--app-green) !important;
  border-color: var(--app-green) !important;
}
.stButton > button, .stFormSubmitButton > button, [data-testid="stDownloadButton"] > button {
  background-color: var(--app-green) !important;
  color: #FFFFFF !important;
  border: 1px solid var(--app-green) !important;
}
.stButton > button:hover, .stFormSubmitButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
  background-color: #2E9E2E !important;
  border-color: #2E9E2E !important;
  color: #FFFFFF !important;
}
/* Force green selected chips in Streamlit's BaseWeb multiselect. */
div[data-baseweb="select"] span[data-baseweb="tag"],
div[data-baseweb="select"] [data-baseweb="tag"],
div[data-baseweb="select"] [data-baseweb="tag"] > span,
div[data-baseweb="select"] [data-baseweb="tag"] > div {
  background: #228B22 !important;
  background-color: #228B22 !important;
  border-color: #228B22 !important;
}
div[data-baseweb="select"] [data-baseweb="tag"] *,
div[data-baseweb="select"] [data-baseweb="tag"] span {
  color: #444444 !important;
  -webkit-text-fill-color: #444444 !important;
}
/* Dropdown popup is rendered in a portal, so target menu/list options globally. */
div[data-baseweb="popover"],
div[data-baseweb="popover"] *,
ul[data-testid*="Dropdown"] ,
ul[data-testid*="Dropdown"] *,
li[role="option"],
div[role="option"] {
  color: #444444 !important;
  -webkit-text-fill-color: #444444 !important;
}
div[data-baseweb="popover"] [role="option"],
li[role="option"],
div[role="option"] {
  background-color: #FFFFFF !important;
}
/* Selected radio indicators: Streamlit's primaryColor is also set in config.toml. */
div[data-testid="stRadio"] label:has(input:checked) > div,
div[data-testid="stRadio"] label:has(input:checked) > div > div,
div[data-testid="stRadio"] label:has(input:checked) > div > div > div,
div[data-testid="stRadio"] [role="radio"][aria-checked="true"] {
  background-color: #228B22 !important;
  border-color: #228B22 !important;
  accent-color: #228B22 !important;
}
div[data-testid="stRadio"] input[type="radio"] {
  accent-color: #228B22 !important;
}
/* Keep dataframe colors as supplied by Streamlit; make column headings bold. */
[data-testid="stDataFrame"] [role="columnheader"],
[data-testid="stDataEditor"] [role="columnheader"] {
  font-weight: 700 !important;
}

</style>
""", unsafe_allow_html=True)

CHART_BAR = "#eab676"
CHART_BACKGROUND = "#76b5c5"

def style_chart(fig):
    fig.update_traces(marker_color=CHART_BAR)
    fig.update_layout(
        paper_bgcolor=CHART_BACKGROUND,
        plot_bgcolor=CHART_BACKGROUND,
        font=dict(color="#FFFFFF"),
        title_font=dict(color="#9DDCFF"),
        legend=dict(font=dict(color="#FFFFFF")),
        xaxis=dict(tickfont=dict(color="#FFFFFF"), title_font=dict(color="#FFFFFF"), gridcolor="rgba(255,255,255,0.30)"),
        yaxis=dict(tickfont=dict(color="#FFFFFF"), title_font=dict(color="#FFFFFF"), gridcolor="rgba(255,255,255,0.30)"),
    )
    return fig
DEFAULT_FILE = Path(__file__).with_name("Amazon_Reviews_Sentiment_Analysis.xlsx")

STOPWORDS = set("a an the and or but if then than so because as at by for from in into of on onto to up with without is are was were be been being am do does did have has had it its this that these those i me my we our you your he she they them their who what when where why how not no nor very really just can could would should will may might must about after before over under again more most less much many some any all each other also there here their them his her our out off get got make made one two use used using service amazon product products customer customers review reviews star stars buy bought order orders delivery delivered shipping like love hate time day days week month price prices".split())

TOPICS = {
    "Delivery & drivers": ["deliver", "delivery", "driver", "courier", "parcel", "package", "late", "tracking", "doorstep", "left outside", "missed delivery", "van"],
    "Customer service & support": ["customer service", "support", "agent", "representative", "complaint", "contact", "call centre", "rude", "helpful", "help desk", "response"],
    "Returns, refunds & exchanges": ["refund", "return", "exchange", "reimburse", "replacement", "return label", "money back", "cancelled order"],
    "Prime membership & pricing": ["prime", "membership", "subscription", "fee", "price", "expensive", "cost", "value for money", "worth it", "increase in price"],
    "Product quality & condition": ["faulty", "defect", "broken", "damaged", "poor quality", "quality", "fake", "counterfeit", "not as described", "missing parts", "stopped working"],
    "Marketplace & third-party sellers": ["seller", "third party", "marketplace", "vendor", "fraudulent seller", "scam"],
    "Account, billing & payment": ["account", "charged", "charge", "billing", "payment", "refund", "unauthorised", "security", "locked out", "password"],
    "Website, app & ordering": ["website", "app", "application", "checkout", "technical", "error message", "glitch", "crash", "order process", "search function"],
    "Product availability & selection": ["out of stock", "unavailable", "availability", "selection", "range", "stock", "discontinued"],
    "Packaging & sustainability": ["packaging", "packaged", "box", "plastic", "waste", "overpackaged", "recyclable"],
    "Communication & policies": ["policy", "terms", "communication", "email", "notification", "misleading", "transparency", "policy change"],
    "Positive experience & praise": ["excellent", "great", "brilliant", "fantastic", "love", "happy", "pleased", "satisfied", "recommend", "convenient", "easy to use", "resolved", "thank you", "thanks", "helpful", "quick delivery"]
}
POSITIVE_CATS = {"Happiness / Joy", "Gratitude / Admiration"}
NEGATIVE_CATS = {"High-intensity hostility, raised voice, or demands for immediate escalation", "Frustration", "Disappointment", "Disapproval / Dissatisfaction", "Worry / Concern"}

@st.cache_data(show_spinner=False)
def load_data(uploaded_bytes=None):
    source = uploaded_bytes if uploaded_bytes is not None else DEFAULT_FILE
    xls = pd.ExcelFile(source)
    sheet = "Labeled Reviews" if "Labeled Reviews" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(source, sheet_name=sheet)
    for c in ["Country", "Review Text", "Review Title", "Sentiment Category", "Rating", "Review Date"]:
        if c not in df.columns: df[c] = ""
    df["Country"] = df["Country"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    df["Review Text"] = df["Review Text"].fillna("").astype(str)
    df["Review Title"] = df["Review Title"].fillna("").astype(str)
    df["Sentiment Category"] = df["Sentiment Category"].fillna("Unlabelled").astype(str).str.strip()
    df["Rating Number"] = pd.to_numeric(df["Rating"].astype(str).str.extract(r"([1-5])")[0], errors="coerce")
    df["Review Date Parsed"] = pd.to_datetime(df["Review Date"], errors="coerce", utc=True)
    df["Review Year"] = df["Review Date Parsed"].dt.year.astype("Int64").astype(str).replace("<NA>", "Unknown")
    df["Full Text"] = (df["Review Title"].str.strip() + ". " + df["Review Text"].str.strip()).str.strip(" .")
    return df

def topic_counts(data):
    counts = Counter()
    examples = {}
    for _, row in data.iterrows():
        text = str(row["Full Text"]).lower()
        for topic, terms in TOPICS.items():
            if any(term in text for term in terms):
                counts[topic] += 1
                examples.setdefault(topic, str(row["Full Text"])[:260])
    return counts, examples

def keywords(data, n=20):
    text = " ".join(data["Full Text"].astype(str).tolist()).lower()
    words = re.findall(r"[a-z][a-z'-]{2,}", text)
    counts = Counter(w for w in words if w not in STOPWORDS and len(w) > 2 and not w.isnumeric())
    return counts.most_common(n)

def make_summary(data, counts):
    if data.empty: return "No reviews match the selected filters."
    if not counts: return f"The {len(data):,} matching reviews do not contain enough recognised topic keywords to identify a dominant issue or compliment reliably."
    ranked = counts.most_common(3)
    cats = data["Sentiment Category"].value_counts()
    dominant_sentiment = cats.index[0] if len(cats) else "mixed sentiment"
    top_topics = [name.lower() for name, _ in ranked]
    positive_share = data["Sentiment Category"].isin(POSITIVE_CATS).mean()
    negative_share = data["Sentiment Category"].isin(NEGATIVE_CATS).mean()
    if positive_share >= 0.55:
        framing = "Customers mainly praise" 
    elif negative_share >= 0.55:
        framing = "Customers mainly report concerns about"
    else:
        framing = "The feedback is mixed, with recurring comments about"
    topic_phrase = ", ".join(top_topics[:-1]) + (" and " if len(top_topics) > 1 else "") + top_topics[-1]
    return f"Across {len(data):,} matching reviews, {framing} {topic_phrase}; the most common assigned sentiment is {dominant_sentiment.lower()}."

st.title("Amazon Reviews Analytics")
st.caption("Explore review patterns using multi-select filters. Select several values within a filter to include any of them; different filters are combined.")
with st.sidebar:
    st.header("Data source")
    uploaded = st.file_uploader("Use a replacement Excel workbook (optional)", type=["xlsx"])
    if uploaded:
        df = load_data(uploaded.getvalue())
    elif DEFAULT_FILE.exists():
        df = load_data()
    else:
        st.error("The workbook was not found beside app.py. Please upload the Excel file in this sidebar.")
        st.stop()
    st.divider()
    st.caption("Filters apply to the Review Year, Rating, Country and Sentiment Category fields.")

page = st.radio("Page", ["Home", "Search filters", "Search output"], horizontal=True)
if page == "Home":
    st.subheader("Dataset overview")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total reviews", f"{len(df):,}")
    c2.metric("Average rating", f"{df['Rating Number'].mean():.2f} / 5" if df["Rating Number"].notna().any() else "N/A")
    c3.metric("Countries", f"{df['Country'].nunique():,}")
    c4.metric("Review years", f"{df['Review Year'].nunique():,}")
    left,right = st.columns(2)
    with left:
        st.markdown("#### Sentiment distribution")
        s = df["Sentiment Category"].value_counts().rename_axis("Sentiment").reset_index(name="Reviews")
        st.plotly_chart(style_chart(px.bar(s, x="Sentiment", y="Reviews", title="Reviews by sentiment")), use_container_width=True)
    with right:
        st.markdown("#### Reviews by year")
        y = df["Review Year"].value_counts().sort_index().rename_axis("Year").reset_index(name="Reviews")
        st.plotly_chart(style_chart(px.bar(y, x="Year", y="Reviews", title="Reviews by year")), use_container_width=True)
    st.info("Open **Search filters** to select multiple years, ratings, countries or sentiment categories. Then choose **Search output** to view results.")
elif page == "Search filters":
    st.subheader("Search filters")
    st.write("Choose one or more values from any filter. Leave a filter empty to include all values for that field.")
    with st.form("review_filters_form"):
        a,b = st.columns(2)
        years = sorted(df["Review Year"].dropna().unique().tolist())
        ratings = sorted([int(x) for x in df["Rating Number"].dropna().unique().tolist()])
        countries = sorted(df["Country"].dropna().unique().tolist())
        sentiments = sorted(df["Sentiment Category"].dropna().unique().tolist())
        selected_years = a.multiselect("Review Year", years, placeholder="All years", key="years")
        selected_ratings = b.multiselect("Rating (stars)", ratings, format_func=lambda x: f"{x} star" if x == 1 else f"{x} stars", placeholder="All ratings", key="ratings")
        selected_countries = a.multiselect("Country", countries, placeholder="All countries", key="countries")
        selected_sentiments = b.multiselect("Sentiment Category", sentiments, placeholder="All sentiment categories", key="sentiments")
        go = st.form_submit_button("Go — apply filters", type="primary")
    if go:
        st.session_state["review_filter_values"] = {"years": selected_years, "ratings": selected_ratings, "countries": selected_countries, "sentiments": selected_sentiments}
        st.session_state["search_done"] = True
        st.success("Filters applied. Open Search output to view the analysis.")
    st.caption("Tip: click a selected value's × to remove it. Multiple values in the same dropdown are included with OR logic.")
else:
    st.subheader("Search output")
    f = st.session_state.get("review_filter_values", {"years": [], "ratings": [], "countries": [], "sentiments": []})
    result = df.copy()
    if f["years"]: result = result[result["Review Year"].isin(f["years"])]
    if f["ratings"]: result = result[result["Rating Number"].isin(f["ratings"])]
    if f["countries"]: result = result[result["Country"].isin(f["countries"])]
    if f["sentiments"]: result = result[result["Sentiment Category"].isin(f["sentiments"])]
    st.markdown("**Selected filters**")
    labels = [("Review Year", f["years"]), ("Rating", [f"{x} star(s)" for x in f["ratings"]]), ("Country", f["countries"]), ("Sentiment Category", f["sentiments"])]
    st.write(" | ".join(f"**{k}:** {', '.join(map(str,v)) if v else 'All'}" for k,v in labels))
    m1,m2,m3 = st.columns(3)
    m1.metric("Matching reviews", f"{len(result):,}")
    m2.metric("Average rating", f"{result['Rating Number'].mean():.2f} / 5" if result["Rating Number"].notna().any() else "N/A")
    m3.metric("Countries represented", f"{result['Country'].nunique():,}")
    if result.empty:
        st.warning("No reviews match these filters. Return to Search filters and broaden your selections.")
    else:
        kw = keywords(result)
        topics, examples = topic_counts(result)
        st.markdown("### Main problem or compliment areas")
        st.info(make_summary(result, topics))
        st.caption("Topic counts are keyword-rule estimates; one review can match more than one topic, and ambiguous wording may be misclassified.")
        left,right = st.columns(2)
        with left:
            st.markdown("### Top 20 keywords")
            if kw:
                kdf = pd.DataFrame(kw, columns=["Keyword", "Occurrences"])
                st.dataframe(kdf, use_container_width=True, hide_index=True)
            else: st.write("No meaningful keywords found.")
        with right:
            st.markdown("### Reviews by country")
            country_df = result["Country"].value_counts().rename_axis("Country").reset_index(name="Reviews")
            country_df["Share"] = (country_df["Reviews"] / len(result) * 100).round(1)
            st.plotly_chart(style_chart(px.bar(country_df.head(25), x="Country", y="Reviews", title="Top countries (up to 25)")), use_container_width=True)
            st.dataframe(country_df, use_container_width=True, hide_index=True)
        st.markdown("### Topic breakdown")
        if topics:
            tdf = pd.DataFrame([(k,v,examples.get(k,"")) for k,v in topics.most_common()], columns=["Topic", "Reviews mentioning topic", "Example review excerpt"])
            st.plotly_chart(style_chart(px.bar(tdf, x="Reviews mentioning topic", y="Topic", orientation="h", title="Recognised problem and compliment themes")), use_container_width=True)
            st.dataframe(tdf, use_container_width=True, hide_index=True)
        st.markdown("### Matching reviews")
        show_cols = [c for c in ["Review Date", "Country", "Rating", "Sentiment Category", "Review Title", "Review Text"] if c in result.columns]
        st.dataframe(result[show_cols], use_container_width=True, hide_index=True)
        st.download_button("Download filtered reviews as CSV", result[show_cols].to_csv(index=False).encode("utf-8-sig"), file_name="filtered_amazon_reviews.csv", mime="text/csv")
