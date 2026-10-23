import io
import re
import sys
import tempfile
from pathlib import Path
import requests
import streamlit as st
from PIL import Image

from make_poster import build_poster, extract_palette, pick_text_colors

st.set_page_config(page_title="Album Poster Generator", layout="centered")
st.title("🎵 Album Poster Generator")

# --- STEP 1: SEARCH ---
query = st.text_input("Search Artist (or Artist & Album)", value="Pink Floyd", help="e.g., 'טונה', 'Pink Floyd', or 'Drake Views'")

if st.button("Search Albums", type="primary"):
    with st.spinner("Fetching from iTunes..."):
        url = "https://itunes.apple.com/search"
        # Fetch up to 30 results to give a good dropdown list
        res = requests.get(url, params={"term": query, "entity": "album", "limit": 30}).json()
        
        if not res.get("results"):
            st.error("No albums found. Try checking your spelling!")
            st.session_state.pop("album_results", None)
        else:
            # Filter out singles (trackCount <= 1) and karaoke/tribute albums
            valid_albums = []
            for a in res["results"]:
                if a.get("trackCount", 0) > 1 and "Karaoke" not in a.get("collectionName", ""):
                    valid_albums.append(a)
            
            if not valid_albums:
                st.error("Only singles found. Try another search.")
            else:
                st.session_state["album_results"] = valid_albums
                st.session_state["selected_index"] = 0

# --- STEP 2: SELECT FROM DROPDOWN ---
if "album_results" in st.session_state:
    albums = st.session_state["album_results"]
    album_options = [f"{a['artistName']} - {a['collectionName']} ({a.get('releaseDate', '')[:4]})" for a in albums]
    
    st.divider()
    selected_option = st.selectbox("Select the correct album:", album_options)
    selected_album = albums[album_options.index(selected_option)]

    if st.button("Preview & Load Album"):
        with st.spinner("Loading high-res cover and tracklist..."):
            album_id = selected_album["collectionId"]
            track_res = requests.get("https://itunes.apple.com/lookup", params={"id": album_id, "entity": "song"}).json()
            tracks = [item["trackName"] for item in track_res.get("results", []) if item.get("wrapperType") == "track"]

            st.session_state["album_data"] = {
                "artist": selected_album["artistName"],
                "title": selected_album["collectionName"],
                "date": selected_album.get("releaseDate", "")[:10],
                "genres": selected_album.get("primaryGenreName", "").replace("/", " | "),
                "cover_url": selected_album["artworkUrl100"].replace("100x100bb", "1000x1000bb"),
                "tracks": tracks
            }

            img_data = requests.get(st.session_state["album_data"]["cover_url"]).content
            temp_path = Path(tempfile.gettempdir()) / "temp_cover.jpg"
            temp_path.write_bytes(img_data)
            
            img = Image.open(temp_path)
            st.session_state["palette"] = extract_palette(img)
            st.session_state["cover_path"] = str(temp_path)

# --- STEP 3: PREVIEW & CUSTOMIZE ---
if "album_data" in st.session_state:
    data = st.session_state["album_data"]
    st.divider()
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.image(data["cover_url"], use_container_width=True)
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

    # --- STEP 4: GENERATE PPTX ---
    if st.button("🚀 Generate PPTX Poster"):
        filename = re.sub(r'[\\/:*?"<>|]', "", f"{data['artist']} - {data['title']}") + ".pptx"
        out_path = Path(tempfile.gettempdir()) / filename

        def custom_extract_palette(img):
            return updated_palette

        import make_poster
        original_func = make_poster.extract_palette
        make_poster.extract_palette = custom_extract_palette
        
        build_poster(
            cover_path=st.session_state["cover_path"],
            artist=data["artist"],
            title=data["title"],
            date=data["date"],
            genres=data["genres"],
            tracks=data["tracks"],
            out_path=out_path
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