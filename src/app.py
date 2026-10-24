import io
import re
import sys
import tempfile
from pathlib import Path
import requests
import streamlit as st
from PIL import Image

from make_poster import build_poster, extract_palette, pick_text_colors

st.set_page_config(page_title="Album Poster Generator", layout="wide")

# --- SIDEBAR: GLOBAL SETTINGS ---
with st.sidebar:
    st.header("⚙️ Poster Settings")
    poster_size = st.selectbox("Canvas Size", ["A4 (8.3 x 11.7)", "US Letter (8.5 x 11)", "A3 (11.7 x 16.5)", "Poster (18 x 24)"])
    poster_theme = st.selectbox("Design Theme", ["Light - white background", "Dark - black background"])
    st.caption("Note: Size and Theme engine will be connected in the next update!")
    st.divider()
    st.markdown("Built with Python & Streamlit")

st.title("🎵 Album Poster Generator")

# --- MAIN UI: TABS ---
tab1, tab2 = st.tabs(["🔍 Search Database", "💿 Custom Mixtape"])

# --- TAB 1: SEARCH ITUNES ---
with tab1:
    query = st.text_input("Search Artist (or Artist & Album)", value="Pink Floyd", help="e.g., 'טונה', 'Pink Floyd'")
    
    if st.button("Search Albums", type="primary"):
        with st.spinner("Fetching from iTunes..."):
            url = "https://itunes.apple.com/search"
            res = requests.get(url, params={"term": query, "entity": "album", "limit": 30}).json()
            
            if not res.get("results"):
                st.error("No albums found. Try checking your spelling!")
                st.session_state.pop("album_results", None)
            else:
                valid_albums = [a for a in res["results"] if a.get("trackCount", 0) > 1 and "Karaoke" not in a.get("collectionName", "")]
                if not valid_albums:
                    st.error("Only singles found. Try another search.")
                else:
                    st.session_state["album_results"] = valid_albums

    if "album_results" in st.session_state:
        albums = st.session_state["album_results"]
        album_options = [f"{a['artistName']} - {a['collectionName']} ({a.get('releaseDate', '')[:4]})" for a in albums]
        
        selected_option = st.selectbox("Select the correct album:", album_options)
        selected_album = albums[album_options.index(selected_option)]
    
        if st.button("Load Album Data"):
            with st.spinner("Loading high-res cover and tracklist..."):
                album_id = selected_album["collectionId"]
                track_res = requests.get("https://itunes.apple.com/lookup", params={"id": album_id, "entity": "song"}).json()
                tracks = [item["trackName"] for item in track_res.get("results", []) if item.get("wrapperType") == "track"]
    
                st.session_state["album_data"] = {
                    "artist": selected_album["artistName"],
                    "title": selected_album["collectionName"],
                    "date": selected_album.get("releaseDate", "")[:10],
                    "genres": selected_album.get("primaryGenreName", "").replace("/", " | "),
                    "tracks": tracks
                }
    
                img_data = requests.get(selected_album["artworkUrl100"].replace("100x100bb", "1000x1000bb")).content
                temp_path = Path(tempfile.gettempdir()) / "temp_cover.jpg"
                temp_path.write_bytes(img_data)
                
                img = Image.open(temp_path)
                st.session_state["palette"] = extract_palette(img)
                st.session_state["cover_path"] = str(temp_path)

# --- TAB 2: CUSTOM UPLOAD ---
with tab2:
    st.subheader("Build a Custom Poster")
    colA, colB = st.columns(2)
    with colA:
        custom_artist = st.text_input("Artist Name", "My Band")
        custom_title = st.text_input("Album Title", "The Mixtape")
        custom_date = st.text_input("Release Date", "2024-01-01")
        custom_genres = st.text_input("Genres", "Indie | Alt-Pop")
    with colB:
        custom_tracks = st.text_area("Tracklist (One song per line)", "Intro\nFirst Song\nThe Hit\nOutro", height=130)
        custom_image = st.file_uploader("Upload Cover Art", type=["png", "jpg", "jpeg", "webp"])

    if st.button("Load Custom Data", type="primary") and custom_image:
        temp_path = Path(tempfile.gettempdir()) / "custom_temp_cover.jpg"
        temp_path.write_bytes(custom_image.getvalue())
        
        img = Image.open(temp_path)
        st.session_state["palette"] = extract_palette(img)
        st.session_state["cover_path"] = str(temp_path)
        
        st.session_state["album_data"] = {
            "artist": custom_artist,
            "title": custom_title,
            "date": custom_date,
            "genres": custom_genres,
            "tracks": [t.strip() for t in custom_tracks.split("\n") if t.strip()]
        }

# --- STEP 3: PREVIEW & GENERATE (Shared by both tabs) ---
if "album_data" in st.session_state:
    data = st.session_state["album_data"]
    st.divider()
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.image(st.session_state["cover_path"], use_container_width=True)
    with col2:
        st.subheader(data["title"])
        st.write(f"**Artist:** {data['artist']}")
        st.write(f"**Release Date:** {data['date']}")
        st.write(f"**Genres:** {data['genres']}")
        st.write(f"**Total Tracks:** {len(data['tracks'])}")

    st.subheader("🎨 Color Palette")
    cols = st.columns(5)
    updated_palette = []
    for i in range(5):
        with cols[i]:
            hex_val = f"#{st.session_state['palette'][i]}"
            new_color = st.color_picker(f"Color {i+1}", value=hex_val, key=f"color_{i}")
            updated_palette.append(new_color.lstrip("#").upper())

    if st.button("🚀 Generate PPTX Poster"):
        filename = re.sub(r'[\\/:*?"<>|]', "", f"{data['artist']} - {data['title']}") + ".pptx"
        out_path = Path(tempfile.gettempdir()) / filename

        def custom_extract_palette(img):
            return updated_palette

        import make_poster
        original_func = make_poster.extract_palette
        make_poster.extract_palette = custom_extract_palette
        
        # In Step 2, we will pass poster_size and poster_theme to this function!
        build_poster(
            cover_path=st.session_state["cover_path"],
            artist=data["artist"],
            title=data["title"],
            date=data["date"],
            genres=data["genres"],
            tracks=data["tracks"],
            out_path=out_path,
            size_name=poster_size,
            theme_name=poster_theme
        )
        
        make_poster.extract_palette = original_func 

        st.success("Poster generated successfully!")
        with open(out_path, "rb") as f:
            st.download_button(
                label="📥 Download PowerPoint File (.pptx)",
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )