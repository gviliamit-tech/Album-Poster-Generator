import io
import re
import sys
import tempfile
from pathlib import Path
import requests
import streamlit as st
from PIL import Image

# Import layout and color engine from your core script
from make_poster import build_poster, extract_palette, pick_text_colors

st.set_page_config(page_title="Album Poster Generator", layout="centered")
st.title("🎵 Album Poster Generator")
st.markdown("Search for any album, preview/adjust extracted colors, and generate editable `.pptx` posters.")

# --- STEP 1: SEARCH ---
query = st.text_input("Enter Album & Artist", value="Drake - Views", help="Format: 'Artist - Album' works best")

if st.button("Search Album", type="primary"):
    with st.spinner("Fetching album data from iTunes..."):
        url = "https://itunes.apple.com/search"
        res = requests.get(url, params={"term": query, "entity": "album", "limit": 10}).json()
        
        if not res.get("results"):
            st.error("No albums found. Try checking your spelling!")
        else:
            best_match = res["results"][0]
            if " - " in query:
                target_artist = query.split(" - ")[0].strip().lower()
                for item in res["results"]:
                    if target_artist in item.get("artistName", "").lower():
                        best_match = item
                        break
            
            # Fetch Tracklist
            album_id = best_match["collectionId"]
            track_res = requests.get("https://itunes.apple.com/lookup", params={"id": album_id, "entity": "song"}).json()
            tracks = [item["trackName"] for item in track_res.get("results", []) if item.get("wrapperType") == "track"]

            # Save search results in session state
            st.session_state["album_data"] = {
                "artist": best_match["artistName"],
                "title": best_match["collectionName"],
                "date": best_match.get("releaseDate", "")[:10],
                "genres": best_match.get("primaryGenreName", "").replace("/", " | "),
                "cover_url": best_match["artworkUrl100"].replace("100x100bb", "1000x1000bb"),
                "tracks": tracks
            }

            # Download Image to temp path
            img_data = requests.get(st.session_state["album_data"]["cover_url"]).content
            temp_path = Path(tempfile.gettempdir()) / "temp_cover.jpg"
            temp_path.write_bytes(img_data)
            
            # Extract default palette
            img = Image.open(temp_path)
            st.session_state["palette"] = extract_palette(img)
            st.session_state["cover_path"] = str(temp_path)

# --- STEP 2: PREVIEW & CUSTOMIZE ---
if "album_data" in st.session_state:
    data = st.session_state["album_data"]
    st.divider()
    st.subheader(f"Results: {data['artist']} - {data['title']}")

    col1, col2 = st.columns([1, 2])
    with col1:
        st.image(data["cover_url"], use_container_width=True)
    
    with col2:
        st.write(f"**Release Date:** {data['date']}")
        st.write(f"**Genres:** {data['genres']}")
        st.write(f"**Total Tracks:** {len(data['tracks'])}")

    st.subheader("🎨 Extracted Color Palette")
    st.caption("Click to tweak color choices before generating:")
    
    cols = st.columns(5)
    updated_palette = []
    for i in range(5):
        with cols[i]:
            hex_val = f"#{st.session_state['palette'][i]}"
            new_color = st.color_picker(f"Color {i+1}", value=hex_val, key=f"color_{i}")
            updated_palette.append(new_color.lstrip("#").upper())

    # --- STEP 3: GENERATE PPTX ---
    if st.button("🚀 Generate PPTX Poster"):
        filename = re.sub(r'[\\/:*?"<>|]', "", f"{data['artist']} - {data['title']}") + ".pptx"
        out_path = Path(tempfile.gettempdir()) / filename

        # Temporary override palette function to use user-chosen colors
        def custom_extract_palette(img):
            return updated_palette

        # Generate Poster
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
        
        make_poster.extract_palette = original_func # Reset

        st.success("Poster generated successfully!")
        
        with open(out_path, "rb") as f:
            st.download_button(
                label="📥 Download PowerPoint File (.pptx)",
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )