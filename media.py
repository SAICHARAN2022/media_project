import os
import pandas as pd
from googleapiclient.discovery import build
import yt_dlp

# YouTube API setup
API_KEY = os.getenv("YOUTUBE_API_KEY")





  # Replace with your YouTube Data API key
YOUTUBE_API_SERVICE_NAME = "youtube"
YOUTUBE_API_VERSION = "v3"



def fetch_video_data(video_url):
    """
    Fetches YouTube video metadata using yt-dlp.
    """
    try:
        ydl_opts = {}
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(video_url, download=False)
            video_data = {
                "Video ID": info.get("id"),
                "Title": info.get("title"),
                "Description": info.get("description"),
                "Published Date": info.get("upload_date"),  # Format: YYYYMMDD
                "View Count": info.get("view_count"),
                "Like Count": info.get("like_count"),  # Can be None if unavailable
                "Comment Count": info.get("comment_count"),  # Can be None if unavailable
                "Duration (seconds)": info.get("duration"),
                "Thumbnail URL": info.get("thumbnail"),
            }
            return video_data
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def fetch_video_comments(video_id, max_comments=100):
    """
    Fetches the latest comments and their replies from a YouTube video.
    """
    try:
        youtube = build(YOUTUBE_API_SERVICE_NAME, YOUTUBE_API_VERSION, developerKey=API_KEY)
        comments = []
        next_page_token = None  # Initialize pagination

        while True:
            request = youtube.commentThreads().list(
                part="snippet",
                videoId=video_id,
                maxResults=min(max_comments, 100),  # Fetch up to 100 comments per request
                order="time",  # Fetch latest comments first
                pageToken=next_page_token
            )
            response = request.execute()

            if not response.get("items"):
                print("No comments found for the video.")
                break

            # Iterate through the comments
            for item in response.get("items", []):
                try:
                    top_comment = item["snippet"]["topLevelComment"]["snippet"]
                    comment_id = item["snippet"]["topLevelComment"].get("id", "No ID")
                    author = top_comment["authorDisplayName"]
                    text = top_comment["textOriginal"]
                    published_at = top_comment["publishedAt"]
                    like_count = top_comment["likeCount"]
                    parent_comment_id = ""  # Top-level comment doesn't have a parent
                
                    # Prepare the top-level comment data
                    comments.append({
                        "Video ID": video_id,
                        "Comment ID": comment_id,
                        "Comment Text": text,
                        "Author Name": author,
                        "Published Date": published_at,
                        "Like Count": like_count,
                        "Reply To": parent_comment_id
                    })
                    
                    # Now fetch the replies for this comment (if any)
                    replies_request = youtube.comments().list(
                        part="snippet",
                        parentId=comment_id,
                        maxResults=100  # You can adjust this to fetch more replies if needed
                    )
                    replies_response = replies_request.execute()
                    
                    # If there are replies, add them to the comments list
                    for reply_item in replies_response.get("items", []):
                        reply = reply_item["snippet"]
                        reply_comment_id = reply.get("id", "No ID")
                        reply_author = reply["authorDisplayName"]
                        reply_text = reply["textOriginal"]
                        reply_published_at = reply["publishedAt"]
                        reply_like_count = reply["likeCount"]
                        
                        # Add reply data to the list with the parent comment ID
                        comments.append({
                            "Video ID": video_id,
                            "Comment ID": reply_comment_id,
                            "Comment Text": reply_text,
                            "Author Name": reply_author,
                            "Published Date": reply_published_at,
                            "Like Count": reply_like_count,
                            "Reply To": comment_id  # Set parent comment ID
                        })
                except KeyError as e:
                    print(f"KeyError: Missing key in comment data: {e}")
                    continue  # Skip this item and continue with the next comment

            # Check if there is another page of comments
            next_page_token = response.get("nextPageToken")
            if not next_page_token:
                break  # Exit loop if there are no more pages

        return comments
    except Exception as e:
        print(f"An error occurred while fetching comments: {e}")
        return []

# Replace with your YouTube video URL
VIDEO_URL = "https://www.youtube.com/watch?v=gE3rXcUjqo4"

# Fetch video data
video_data = fetch_video_data(VIDEO_URL)

if video_data:
    video_id = video_data["Video ID"]
    
    comments = fetch_video_comments(video_id)

    # Save to Excel
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = os.path.join(downloads_path, "youtube_video_data.xlsx")

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df_video = pd.DataFrame([video_data])
        df_video.to_excel(writer, sheet_name="Video Metadata", index=False)
        
        df_comments = pd.DataFrame(comments)
        df_comments.to_excel(writer, sheet_name="Latest Comments", index=False)

    print(f"Data saved to: {file_path}")
    print(video_data)
    print(f"Fetched {len(comments)} comments.")
