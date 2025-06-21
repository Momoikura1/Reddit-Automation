import praw
import pandas as pd
import datetime
import os

def extract_post_id(reddit_link):
    try:
        if '/comments/' in reddit_link:
            post_id = reddit_link.split('/comments/')[1].split('/')[0]
            return post_id
        return None
    except Exception:
        return None

def check_user_status(reddit, username):
    try:
        user = reddit.redditor(username)
        try:
            _ = user.id
            try:
                _ = user.created_utc
                return "active"
            except Exception:
                return "cannot be messaged"
        except Exception:
            return "suspended or deleted"
    except Exception:
        return "error"

def get_top_commenters(submission, reddit, top_n=3):
    submission.comments.replace_more(limit=0)
    commenters = []
    for comment in submission.comments:
        if comment.author and comment.author.name != '[deleted]' and not comment.distinguished:
            status = check_user_status(reddit, comment.author.name)
            if status == "active":
                commenters.append((comment.author.name, comment.score, comment))
    commenters = sorted(commenters, key=lambda x: x[1], reverse=True)[:top_n]
    return commenters

def analyze_and_message_post(reddit, post_url, poster_msg, commenter_msg, results=None):
    post_id = extract_post_id(post_url)
    if not post_id:
        print(f"\n❌ Invalid Reddit post link: {post_url}")
        return

    try:
        submission = reddit.submission(id=post_id)
        subreddit_name = f"r/{submission.subreddit.display_name}"

        if getattr(submission, "removed", False):
            post_status = "deleted"
        elif getattr(submission, "archived", False):
            post_status = "archived"
        else:
            post_status = "active"

        if submission.author:
            poster_username = submission.author.name
            poster_status = check_user_status(reddit, poster_username)
            if poster_status == "active":
                poster_status_str = "active"
            elif poster_status == "cannot be messaged":
                poster_status_str = "cannot be messaged"
            elif poster_status == "suspended or deleted":
                try:
                    _ = reddit.redditor(poster_username).id
                    poster_status_str = "suspended"
                except Exception:
                    poster_status_str = "deleted"
            else:
                poster_status_str = "error"
        else:
            poster_username = "[deleted]"
            poster_status_str = "deleted"

        best_commenters = get_top_commenters(submission, reddit, top_n=3)

        print("\n=== Analysis Result ===")
        print(f"Subreddit: {subreddit_name}")
        print(f"Poster: u/{poster_username}")
        print(f"Poster status: {poster_status_str}")
        print(f"Post status: {post_status}")
        if best_commenters:
            print("Top 3 commenters who can be messaged:")
            for i, (username, score, comment_obj) in enumerate(best_commenters, 1):
                print(f"{i}. u/{username} (upvotes: {score})")
        else:
            print("No commenters who can be messaged found.")

        # --- Message Generation ---
        if poster_status_str == "active":
            msg = poster_msg.replace("link of the post", post_url)
            print(f"\nGenerated message for poster u/{poster_username}:")
            print(msg)
        else:
            msg = ""
            print(f"Poster u/{poster_username} cannot be messaged.")

        commenter_msgs = []
        for username, score, comment_obj in best_commenters:
            comment_link = f"https://www.reddit.com{comment_obj.permalink}"
            cmsg = commenter_msg.replace("link of the comment", comment_link)
            print(f"\nGenerated message for commenter u/{username}:")
            print(cmsg)
            commenter_msgs.append({"username": username, "score": score, "comment_link": comment_link, "message": cmsg})

        # --- Save to results list if provided ---
        if results is not None:
            row = {
                "Subreddit": subreddit_name,
                "Post Link": post_url,
                "Poster": f"u/{poster_username}",
                "Poster Status": poster_status_str,
                "Post Status": post_status,
                "Poster Message": msg,
            }
            # Add top 3 commenters info
            for idx in range(3):
                if idx < len(commenter_msgs):
                    row[f"Commenter {idx+1}"] = f"u/{commenter_msgs[idx]['username']}"
                    row[f"Commenter {idx+1} Upvotes"] = commenter_msgs[idx]['score']
                    row[f"Commenter {idx+1} Link"] = commenter_msgs[idx]['comment_link']
                    row[f"Commenter {idx+1} Message"] = commenter_msgs[idx]['message']
                else:
                    row[f"Commenter {idx+1}"] = ""
                    row[f"Commenter {idx+1} Upvotes"] = ""
                    row[f"Commenter {idx+1} Link"] = ""
                    row[f"Commenter {idx+1} Message"] = ""
            results.append(row)

    except Exception as e:
        print(f"❌ Error processing post: {str(e)}")

def process_multiple_links(reddit, poster_msg, commenter_msg):
    import os
    results = []
    # Use your requested fixed filename and path
    filename = r"c:\Users\markd\OneDrive\Desktop\Reddit Automation\reddit_automation_results_20250621_204151.xlsx"
    while True:
        print("\nPaste your Reddit post links (one per line). Enter an empty line to finish:")
        links = []
        while True:
            link = input()
            if not link.strip():
                break
            links.append(link.strip())
        if not links:
            print("No links entered.")
        else:
            for idx, link in enumerate(links, 1):
                print(f"\nProcessing link {idx} of {len(links)}:")
                analyze_and_message_post(reddit, link, poster_msg, commenter_msg, results)
            # Append to Excel file
            if os.path.exists(filename):
                old_df = pd.read_excel(filename)
                new_df = pd.DataFrame(results)
                df = pd.concat([old_df, new_df], ignore_index=True)
            else:
                df = pd.DataFrame(results)
            df.to_excel(filename, index=False)
            print(f"\n✅ Results saved to {filename}")
            results.clear()  # Clear results for the next batch
        # Ask user if they want to add more links or exit
        choice = input("\nDo you want to add more links? (y/n): ").strip().lower()
        if choice != "y":
            print("Exiting program.")
            break

def main():
    print("=== Reddit Automation Setup ===")
    print("To use this program, you need to create a Reddit app for API access.")
    print("Instructions:")
    print("  1. Go to https://www.reddit.com/prefs/apps")
    print("  2. Click 'Create App' or 'Create Another App'")
    print("  3. Choose 'script' as the app type")
    print("  4. Fill in the name and redirect uri (can be http://localhost:8080)")
    print("  5. After creation, copy your client_id (under the app name) and client_secret")
    print("  6. Use your Reddit username and password for login below")
    print()
    client_id = input("Enter your Reddit client_id: ").strip()
    client_secret = input("Enter your Reddit client_secret: ").strip()
    username = input("Enter your Reddit username: ").strip()
    password = input("Enter your Reddit password: ").strip()
    try:
        reddit = praw.Reddit(
            client_id=client_id,
            client_secret=client_secret,
            username=username,
            password=password,
            user_agent=f"RedditAutomationBot:v1.0 by /u/{username}"
        )
        reddit.user.me()
        print("✅ Successfully logged in!\n")
    except Exception as e:
        print(f"❌ Failed to authenticate: {str(e)}")
        return

    print("\nEnter the message to send to the poster.")
    print('Use "link of the post" in your message to insert the post link.')
    poster_msg = input("Poster message: ").strip()

    print("\nEnter the message to send to the commenters.")
    print('Use "link of the comment" in your message to insert the comment link.')
    commenter_msg = input("Commenter message: ").strip()

    print("\nChoose mode:")
    print("1. Single Reddit post link")
    print("2. Multiple Reddit post links")
    mode = input("Enter 1 or 2: ").strip()
    if mode == "1":
        post_url = input("Enter the Reddit post link: ").strip()
        analyze_and_message_post(reddit, post_url, poster_msg, commenter_msg)
    elif mode == "2":
        process_multiple_links(reddit, poster_msg, commenter_msg)
    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()