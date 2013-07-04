import time
from twython import Twython
from collections import defaultdict

#################################################################################
## io

user = 'jp_trends'
wc = defaultdict(str)

APP_KEY = 'uPmFUwPSHx3bYb40G7OxUw'
APP_SECRET = 'QkKn9YnQ3BZlosjVso6qRbgIb1dd1buCHHWwrc4iNNE'
OAUTH_TOKEN = '489132649-ozR4lvl04kBo1vNZm6FvfFe7v6llZ4L2C8xSnmWG'
OAUTH_TOKEN_SECRET = 'KE9CrNJDmNO9ADxERGEOXQV7tSNBQN0rhVgLXKSts'

#################################################################################
## Function Definitions

# Cleans Japanese Text
def text_scrape(tweet):
    tweet_text = tweet['text'].split(u'\u6d88')[0]
    tweet_text = tweet_text.rsplit(u'\u65b0',1)[0]
    for char in ['[', u'\u65b0']:
        tweet_text = tweet_text.replace(char, '')
    return tweet_text.split(']')

# Counts # of Words
def word_count(word_list, date):
    for tweet in word_list:
        print tweet
    return

#################################################################################
## Main Program

def main():
    t0 = time.clock()
    twitter = Twython(APP_KEY, APP_SECRET, OAUTH_TOKEN, OAUTH_TOKEN_SECRET)

    # Loop through 16 Pages (Twitter API Limit of 3200)
    maxID = None
    for page in range(1,17):
        print page, maxID
        user_timeline = twitter.get_user_timeline(
                            screen_name = user,
                            count = 200,
                            include_rts = 1,
                            exclude_replies = True,
                            max_id = maxID)

        for tweet in user_timeline:
            tweet_text = text_scrape(tweet)
            print 'ORIGINAL:', tweet['text']
            word_count(tweet_text, tweet['created_at'])
            raw_input('...')
            maxID = tweet['id']

    for word in wc:
        print word, wc[word]

    t1 = time.clock()
    return 'Process completed! Duration:', t1-t0



main()


"""

"""