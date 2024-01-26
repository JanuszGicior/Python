import tweepy
import os
import schedule
import time
API_Key =  "Your_API_Key "
API_Secret_Key =  "Your_API_Secret_Key"
Access_Token = "Your_Access_Token"
Access_Secret_Token = "Your_Access_Secret_Token"
Bearer_Token = "Your_Bearer_Token"
client = tweepy.Client(Bearer_Token,API_Key, API_Secret_Key, Access_Token, Access_Secret_Token)


def post_to_twitter():
    """
    Authenticates with Twitter API and creates a Tweepy API object.
    Uploads media and creates a tweet using Twitter API v2.
    Deletes media file after posting (optional).
    """
    # Authenticate and create a Tweepy API object
    auth = tweepy.OAuth1UserHandler(API_Key, API_Secret_Key, Access_Token,Access_Secret_Token)
    #auth1 = tweepy.OAuth2UserHandler(Bearer_Token)
    api = tweepy.API(auth)
    #api1 = tweepy.API(auth1)

    # Upload media and create tweet using Twitter API v2
    
    media = api.media_upload('hustle1.jpg')
    print(media)
    client.create_tweet(text="Hustle", media_ids=[media.media_id])
    #client.like(1722978441895547220)

    # Delete media file after posting (optional)
    #os.remove('image1.jpg')



post_to_twitter()


# Set the date and time for the tweet
tweet_date = '10:42'

# Schedule the tweet at the specified date and time
schedule.every().day.at(tweet_date).do(post_to_twitter)

# Keep running the script to check for scheduled tasks
while True:
    schedule.run_pending()
    time.sleep(1) 
