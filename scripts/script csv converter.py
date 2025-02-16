import os  
from dotenv import load_dotenv
import base64
import pandas as pd
from openai import AzureOpenAI
import tkinter as tk
from tkinter import filedialog

# Load environment variables from .env file
load_dotenv()

# Create and hide the root window for file dialogs
root = tk.Tk()
root.withdraw()

# Initialize Azure OpenAI client with your credentials
endpoint = os.getenv("ENDPOINT_URL", "https://azureopenaiapigiorgio.openai.azure.com/")  
deployment = os.getenv("DEPLOYMENT_NAME", "gpt-4o")  
subscription_key = os.getenv("AZURE_OPENAI_API_KEY")

client = AzureOpenAI(  
    azure_endpoint=endpoint,  
    api_key=subscription_key,  
    api_version="2024-05-01-preview",
)

# Open file dialog for selecting input file
print("Please select your input file...")
input_file = filedialog.askopenfilename(
    title="Select file to process",
    filetypes=[
        ("Excel files", "*.xlsx"),
        ("Excel 97-2003 files", "*.xls"),
        ("CSV files", "*.csv"),
        ("All files", "*.*")
    ]
)

if not input_file:
    print("No file selected. Exiting...")
    exit()

# Read the selected file with appropriate method based on file type
print(f"Reading file: {input_file}")
try:
    if input_file.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(input_file)
        print("Successfully read Excel file!")
    else:
        # Try multiple CSV reading methods if needed
        try:
            df = pd.read_csv(input_file, 
                           encoding='utf-8',
                           quoting=1,
                           escapechar='\\',
                           quotechar='"',
                           engine='python')
        except Exception as e:
            print(f"First attempt failed, trying alternative method...")
            try:
                df = pd.read_csv(input_file, 
                               encoding='utf-8',
                               sep='\t',
                               engine='python')
            except Exception as e2:
                print(f"Second attempt failed, trying final method...")
                df = pd.read_csv(input_file, 
                               encoding='utf-8',
                               sep=',',
                               quoting=3,
                               escapechar='\\',
                               engine='python')

    print(f"Successfully loaded file with {len(df)} rows!")
    
except Exception as e:
    print(f"Error reading file: {str(e)}")
    print("Please ensure your file is properly formatted.")
    exit(1)

# Define the complete chat prompt array exactly as in the playground
chat_prompt = [
    {
        "role": "system",
        "content": [
            {
                "type": "text",
                "text": "Please rewrite and paraphrase the following text to create a well-structured, clean, and readable HTML output optimized for readability and SEO. The output should follow these guidelines:\n1. Use appropriate heading tags (H1, H2, etc.) for titles and subtitles, keeping them concise and focused on the main topics.\n2. Ensure all links are converted into proper <a> tags with relevant title attributes.\n3. Use <p> tags for paragraphs and <br> for line breaks.\n4. If there are lists, use semantic <ul>, <ol>, and <li> tags.\n5. For any emphasized or bold text, use <em> or <strong> tags instead of <i> or <b>.\n6. If there are images, include descriptive alt text.\n7. Ensure the HTML is properly indented and easy to read.\n8. Include a <title> tag with a concise, descriptive page title.\n9. Use a <meta> description tag with a compelling page description for SEO.\n10. Identify and categorize the links present in the text (e.g., official websites, informational resources).\n11. Place the categorized links in a separate section below the title, using the \"Useful Links\" heading.\n12. Convert any inline links within the text into proper <a> tags with relevant href and title attributes.\n13. If the input is plain text, identify and convert any URLs into proper <a> tags.\n14. Convey information concisely, avoiding repetition and non-essential content.\n15. Restructure the content to create a logical flow and improve readability, using subheadings to break up the text.\n16. Use bold text sparingly to highlight key information or phrases.\n17. Paraphrase the text extensively to create original content that effectively communicates the main points.\n18. Use HTML entities for special characters and emojis to ensure proper rendering.\nPlease output only the HTML code, without any explanations or additional text.\n19. Don’t add paragraphs about “introduction” and “conclusion” because they are absolutly useless and they take space without adding value. Repeating conclusing is useless becuase it says something already said before. So be coincise, no final repetitive thoughts.\n\nExample of desired output structure:\n<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n  <meta charset=\"UTF-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n  <title>Concise and Descriptive Page Title</title>\n  <meta name=\"description\" content=\"Compelling page description for SEO\">\n</head>\n<body>\n  <main>\n    <h1>Main Heading</h1>\n\n    <div class=\"link-section\">\n      <h2>Useful Links:</h2>\n      <ul>\n        <li><a href=\"link-url\" target=\"_blank\" rel=\"noopener noreferrer\" title=\"Link description\">Link Text</a></li>\n      </ul>\n    </div>\n\n    <h2>Subheading 1</h2>\n    <p>Concisely paraphrased content focused on key points.</p>\n\n    <h2>Subheading 2</h2>\n    <p>Well-structured and readable content that effectively communicates the main ideas.</p>\n\n    ...\n\n  </main>\n</body>\n</html>"
            }
        ]
    },
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "Shitamachi Tanabata Matsuri: Celebrating the Star Festival in Tokyo's Shitamachi Districts\nKeywords\nShitamachi Tanabata Matsuri, Tokyo, Star Festival, traditional festival, cultural celebration\nIntroduction\nThe Shitamachi Tanabata Matsuri is a vibrant and traditional festival held in various Shitamachi districts of Tokyo. Celebrated annually on July 7th, this festival commemorates the Star Festival and showcases the rich cultural heritage of Japan. Visitors can immerse themselves in the colorful decorations, lively parades, and festive atmosphere that characterize this beloved event.\nOrigin and Significance\nThe Shitamachi Tanabata Matsuri has its roots in the ancient Japanese legend of the Weaver Princess and the Cowherd. According to the legend, the Weaver Princess and Cowherd were only allowed to meet once a year on the seventh day of the seventh lunar month. This festival, celebrated on July 7th, celebrates their reunion and serves as a time for people to make wishes, write them on colorful strips of paper, and hang them on bamboo branches.\nFestive Decorations and Activities\nDuring the Shitamachi Tanabata Matsuri, the streets come alive with vibrant decorations made of colorful paper streamers, known as \"tanzaku,\" and intricately crafted bamboo ornaments. Visitors can stroll through the festival stalls, enjoy traditional street food, and participate in various activities such as writing wishes on tanzaku, trying their hand at traditional games, and watching captivating performances.\nShitamachi Districts and Community Spirit\nThe festival takes place in Tokyo's Shitamachi districts, including Asakusa, Ueno, and Kappabashi. These areas are known for their traditional, nostalgic atmosphere and strong sense of community. The Shitamachi Tanabata Matsuri provides an opportunity for locals and visitors alike to come together, experience the warmth of the community, and celebrate Japanese culture and traditions. Useful Wikipedia Links:\n1. <u>Shitamachi</u>\n2. <u>Tanabata</u> Google Maps Links:\n3. <u>Asakusa</u>\n4. <u>Ueno</u>\n5. <u>Kappabashi</u>"
            }
        ]
    },
    {
        "role": "assistant",
        "content": [
            {
                "type": "text",
                "text": "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n  <meta charset=\"UTF-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n  <title>Shitamachi Tanabata Matsuri</title>\n  <meta name=\"description\" content=\"Explore the vibrant Shitamachi Tanabata Matsuri, a traditional festival held in Tokyo's Shitamachi districts. Immerse yourself in colorful decorations, lively parades, and festive atmosphere.\">\n  <style>\n    .link-section {\n      display: flex;\n      justify-content: space-around;\n      margin-bottom: 20px;\n    }\n    .link-section h2 {\n      margin: 0 0 10px;\n    }\n    .link-section ul {\n      padding-left: 20px;\n    }\n    .link-section .separator {\n      border-left: 1px solid #ccc;\n      margin: 0 20px;\n    }\n    .link-section a {\n      font-size: 18px;\n      text-decoration: underline;\n    }\n  </style>\n</head>\n<body>\n  <main>\n    <h1>Shitamachi Tanabata Matsuri</h1>\n\n    <div class=\"link-section\">\n      <div>\n        <h2>Wikipedia Links:</h2>\n        <ul>\n          <li><a href=\"https://en.wikipedia.org/wiki/Shitamachi\" target=\"_blank\" rel=\"noopener noreferrer\">Shitamachi</a></li>\n          <li><a href=\"https://en.wikipedia.org/wiki/Tanabata\" target=\"_blank\" rel=\"noopener noreferrer\">Tanabata</a></li>\n        </ul>\n      </div>\n      <div class=\"separator\"></div>\n      <div>\n        <h2>Google Maps Links:</h2>\n        <ul>\n          <li><a href=\"https://goo.gl/maps/cEjuX3D7W7XzG5Ua7\" target=\"_blank\" rel=\"noopener noreferrer\">Asakusa</a></li>\n          <li><a href=\"https://goo.gl/maps/8xqUJ8uASmWUcUyX8\" target=\"_blank\" rel=\"noopener noreferrer\">Ueno</a></li>\n          <li><a href=\"https://goo.gl/maps/4XyEwcreJ2s7RvUq5\" target=\"_blank\" rel=\"noopener noreferrer\">Kappabashi</a></li>\n        </ul>\n      </div>\n    </div>\n\n    <p>The <strong>Shitamachi Tanabata Matsuri</strong>, held annually on July 7th, is a vibrant festival in Tokyo's Shitamachi districts that celebrates the Star Festival and Japan's rich cultural heritage.</p>\n\n    <p>The festival's <strong>origin</strong> lies in the ancient legend of the Weaver Princess and the Cowherd, who were only allowed to meet once a year on the seventh day of the seventh lunar month.</p>\n\n    <p>During the festival, streets are adorned with <strong>colorful decorations</strong> like paper streamers (tanzaku) and bamboo ornaments. Visitors can enjoy traditional food, games, and captivating performances.</p>\n\n    <p>The festival takes place in Tokyo's <strong>Shitamachi districts</strong>, including Asakusa, Ueno, and Kappabashi, known for their traditional atmosphere and strong sense of community.</p>\n  </main>\n</body>\n</html>"
            }
        ]
    },
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "Understanding Kimono and Yukata\nKimono and yukata are traditional Japanese clothing that have been worn for centuries. In this article, we will explore the differences between these two garments, how to wear kimono, and the various types of kimonos for men and women.\nThe Difference between Kimono and Yukata\nKimono and yukata are similar in shape, but they have distinct differences. Kimonos are made of silk or other high-quality fabrics and are more formal, while yukatas are made of cotton and are worn in a more casual setting. Yukatas are also often brightly colored and decorated with bold patterns.\nHow to Wear Kimono\nWearing a kimono can be a bit complicated, as there are many layers and specific ways to tie the obi. However, once you get the hang of it, it can be a beautiful and elegant experience. The first step is to put on the undergarments, followed by the kimono itself. Next, the obi is tied tightly around the waist, and accessories such as sandals and a purse are added to complete the outfit.\nMale Kimono\nMen's kimonos are typically simpler in design and color than women's kimonos. They are often worn for formal events such as weddings and funerals. Men's kimonos come in various colors and patterns, but they are generally darker and more subdued than women's kimonos.\nKimonos for Women\nWomen's kimonos are much more elaborate than men's kimonos, with intricate patterns and vibrant colors. Women's kimonos are worn for special occasions such as weddings, tea ceremonies, and coming-of-age ceremonies. They are often paired with elaborate hairstyles and accessories.\nConclusion\nKimono and yukata are beautiful and elegant traditional Japanese clothing that have been worn for centuries. Whether you are wearing one for a formal event or just for fun, they are a unique way to experience Japanese culture and style. With the right accessories and techniques, anyone can look and feel like a true Japanese fashion icon."
            }
        ]
    },
    {
        "role": "assistant",
        "content": [
            {
                "type": "text",
                "text": "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n  <meta charset=\"UTF-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n  <title>Kimono and Yukata: Traditional Japanese Clothing</title>\n  <meta name=\"description\" content=\"Discover the differences between kimono and yukata, how to wear them, and the various types for men and women. Explore the beauty of traditional Japanese clothing.\">\n  <style>\n    .link-section {\n      margin-bottom: 20px;\n    }\n    .link-section h2 {\n      margin: 0 0 10px;\n    }\n    .link-section ul {\n      padding-left: 20px;\n    }\n    .link-section a {\n      font-size: 18px;\n      text-decoration: underline;\n    }\n  </style>\n</head>\n<body>\n  <main>\n    <h1>Kimono and Yukata: Traditional Japanese Clothing</h1>\n\n    <div class=\"link-section\">\n      <h2>Useful Links:</h2>\n      <ul>\n        <li><a href=\"https://en.wikipedia.org/wiki/Kimono\" target=\"_blank\" rel=\"noopener noreferrer\">Kimono - Wikipedia</a></li>\n        <li><a href=\"https://en.wikipedia.org/wiki/Yukata\" target=\"_blank\" rel=\"noopener noreferrer\">Yukata - Wikipedia</a></li>\n      </ul>\n    </div>\n\n    <h2>Kimono vs Yukata</h2>\n    <p><strong>Kimono and yukata are both traditional Japanese garments</strong>, but they have some key differences. Kimonos are made from luxurious fabrics like silk and are worn for formal occasions, while yukatas are casual, cotton garments often adorned with vibrant colors and patterns.</p>\n\n    <h2>Wearing Kimono</h2>\n    <p>Putting on a kimono is an <strong>intricate process</strong> involving multiple layers and a specific obi-tying technique. Once mastered, wearing a kimono becomes a graceful and refined experience. Start with undergarments, followed by the kimono, and secure the obi tightly around the waist. Complete the look with sandals and a matching purse.</p>\n\n    <h2>Men's Kimonos</h2>\n    <p>Men's kimonos are <strong>understated yet sophisticated</strong>, often featuring subdued colors and patterns. They are typically worn for formal events like weddings and funerals.</p>\n\n    <h2>Women's Kimonos</h2>\n    <p>Women's kimonos are <strong>true works of art</strong>, boasting intricate designs and bold hues. These stunning garments are reserved for special occasions such as weddings, tea ceremonies, and coming-of-age celebrations. The kimono is often accompanied by elaborate hairstyles and carefully chosen accessories to create a captivating ensemble.</p>\n\n    <p>Kimono and yukata offer a <strong>unique way to immerse oneself in Japanese culture and style</strong>. With the right knowledge and accessories, anyone can embody the timeless elegance of these traditional garments.</p>\n  </main>\n</body>\n</html>"
            }
        ]
    },
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "https://g.co/kgs/3oHLokY\nhttps://www.pokemon.co.jp/shop/en/pokecen/shibuya/\nOpening HoursMonday - Sunday 10:00 am - 9:00 pm\n*Follows Shibuya PARCO business hours.\nHolidayOpen dailyAddress6th Floor, Shibuya PARCO, 15-1, Udagawa-cho, Shibuya-ku, Tokyo, 150-8377\nokémon Center Shibuya\nMore Pokémon than you can throw a pokéball at\nBy Serena Ogawa\nCommunity writer\nIf you’ve got the “catch them all” fever then Shibuya Parco has to be your next stop. Welcome to Shibuya’s exclusive Pokémon Center, newly opened on November 22nd, 2019.\nShibuya is the 14th Pokémon Center to open in Japan, joining Sapporo, Nagoya, and Osaka Pokémon Centers among others. But this is the first and only place for you to find Pikachu in the heart of Tokyo fashion and entertainment, Shibuya.\nLocated on the 6th floor of Shibuya Parco near the Nintendo Tokyo store, Pokémon Center Shibuya features some truly visionary decorations in its entryway. While most Pokémon Centers opt for white walls and colorful starter Pokémon as their mascots to welcome visitors, Shibuya went for a Pokémon a little more fitting its image—Mewtwo. A life-size Mewtwo floats serenely in its tank and as you sneak past, hoping not to wake the legendary psychic Pokémon, you’ll find the entrance.\nBlack walls and floors and a near-endless sea of Pokémon merchandise welcome you inside. Visitors will notice the popular Pikachu mascot decked out in street graffiti—a Shibuya exclusive. Skateboards and cool streetwear-themed merchandise gives the whole Pokémon Center a decidedly Shibuya feel. This is a nightclub-like Pokémon Center made for Shibuya’s nightlife crowd.\nWith the release of Pokémon Sword and Shield, current merchandise reflects the new generation of Pokémon—as of this article, the starter Pokémon (Grookey, Scorbunny, and Sobble) have their merchandise all over the place. You can find plush stuffed animals, figures, collectible card packs, hats, bags, cushions, plates, engagement rings—basically if you think there ought to be Pokémon merchandise of it, you might just find it. Aside from all the exclusive Shibuya-only merchandise and Pokémon plushies, keep your eyes open for the Nintendo x Pokémon collaboration items thanks to the nearby Nintendo Tokyo store—the first and only official Nintendo store in Japan. You’ll also find the Pokémon Sword and Shield “Rotomi” or Pokémon PC Box; give it a talking to and take its Pokémon quiz or play a game with it. And don’t forget the omiyage (souvenir/gift) before you leave. There are various sets of cookies, senbei, and other Japanese snacks in cute and colorful Pokémon containers.\nThe Pokémon Center Shibuya is most definitely worth a visit for Pokémon fans of any age or gender—with an adult appeal not found in the other Centers. Despite its cool, nightclub appearance though you’ll find plenty of your favorite Pokémon merchandise to spend your Pokédollars on."
            }
        ]
    },
    {
        "role": "assistant",
        "content": [
            {
                "type": "text",
                "text": "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n  <meta charset=\"UTF-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n  <title>Pokémon Center Shibuya: A Hub for Pokémon Fans</title>\n  <meta name=\"description\" content=\"Discover Pokémon Center Shibuya, a unique store in the heart of Tokyo's fashion and entertainment district. Explore exclusive merchandise, Pokémon-themed decorations, and a wide range of products for fans of all ages.\">\n</head>\n<body>\n  <main>\n    <h1>Pokémon Center Shibuya: A Hub for Pokémon Fans</h1>\n\n    <div class=\"link-section\">\n      <h2>Useful Links:</h2>\n      <ul>\n        <li><a href=\"https://g.co/kgs/3oHLokY\" target=\"_blank\" rel=\"noopener noreferrer\" title=\"Pokémon Center Shibuya on Google Maps\">Google Maps Location</a></li>\n        <li><a href=\"https://www.pokemon.co.jp/shop/en/pokecen/shibuya/\" target=\"_blank\" rel=\"noopener noreferrer\" title=\"Official Website of Pokémon Center Shibuya\">Official Website</a></li>\n      </ul>\n    </div>\n\n    <p>If you've got the \"catch them all\" fever, then <strong>Pokémon Center Shibuya</strong> is a must-visit destination. Newly opened on November 22nd, 2019, this exclusive store is located on the 6th floor of Shibuya Parco, near the Nintendo Tokyo store.</p>\n\n    <h2>A Unique Pokémon Experience</h2>\n    <p>Pokémon Center Shibuya stands out from other Pokémon Centers with its <strong>visionary decorations</strong> and Shibuya-inspired theme. The entryway features a life-size Mewtwo floating serenely in its tank, setting the tone for a truly immersive experience.</p>\n\n    <h2>Exclusive Merchandise and More</h2>\n    <p>Inside, visitors are greeted by black walls, floors, and a vast array of <strong>Pokémon merchandise</strong>. The popular Pikachu mascot, decked out in street graffiti, is a Shibuya exclusive. From plush stuffed animals and figures to collectible card packs and themed accessories, there's something for every Pokémon fan.</p>\n\n    <h2>Pokémon Sword and Shield Spotlight</h2>\n    <p>With the release of Pokémon Sword and Shield, current merchandise reflects the <strong>new generation of Pokémon</strong>. Don't miss the \"Rotomi\" or Pokémon PC Box, where you can take a Pokémon quiz or play a game. And be sure to check out the Nintendo x Pokémon collaboration items, thanks to the nearby Nintendo Tokyo store.</p>\n\n    <h2>A Must-Visit for Pokémon Fans</h2>\n    <p>Pokémon Center Shibuya offers a <strong>unique appeal</strong> for Pokémon fans of all ages and genders, with its cool, nightclub-like appearance. Before you leave, grab some omiyage (souvenirs) from the selection of cookies, senbei, and other Japanese snacks in cute Pokémon containers.</p>\n\n    <p>Opening Hours: Monday - Sunday, 10:00 am - 9:00 pm<br>\n    Holiday: Open daily<br>\n    Address: 6th Floor, Shibuya PARCO, 15-1, Udagawa-cho, Shibuya-ku, Tokyo, 150-8377</p>\n  </main>\n</body>\n</html>"
            }
        ]
    },
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "https://www.japan.travel/en/destinations/kanto/chiba/\nhttps://www.visitchiba.jp/\n\n# Chiba Prefecture\n\nNext to Tokyo & Home of Narita Airport\n\nVisit Chiba is the official sightseeing, tourism, and tourist attraction information website of Chiba Prefecture, Japan. Chiba is the home of **Narita International Airport** and is conveniently located **next to Tokyo** on a peninsula along Japan’s Pacific coast. Chiba's extensive train and highway bus networks allow visitors to access the rich nature, hidden countryside gems, historic districts, local temples and shrines, seaside hot spring resorts, and world-famous surf breaks found here near Tokyo. Experience Japanese food culture and enjoy fresh seafood too, as Chiba’s mild climate and unique geography make it a hub of agriculture and fishing. We hope to see you soon!\n\n### Don’t Miss\n\n- The popular attractions at Tokyo Disneyland and DisneySea\n- Swimming, surfing and diving at Chiba's beaches\n- Strolling along the historic canals of picturesque Sawara\n- Soaking up the beauty of Shinshoji Temple in Narita, listed in the Michelin Green Guide Japan"
            }
        ]
    },
    {
        "role": "assistant",
        "content": [
            {
                "type": "text",
                "text": "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n  <meta charset=\"UTF-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n  <title>Chiba Prefecture: A Convenient Gateway to Japan</title>\n  <meta name=\"description\" content=\"Discover Chiba Prefecture, located next to Tokyo and home to Narita International Airport. Explore its rich nature, hidden countryside gems, historic districts, and popular attractions like Tokyo Disneyland and DisneySea.\">\n</head>\n<body>\n  <main>\n    <h1>Chiba Prefecture: A Convenient Gateway to Japan</h1>\n\n    <div class=\"link-section\">\n      <h2>Useful Links:</h2>\n      <ul>\n        <li><a href=\"https://www.japan.travel/en/destinations/kanto/chiba/\" target=\"_blank\" rel=\"noopener noreferrer\" title=\"Chiba Prefecture on Japan Travel\">Chiba Prefecture on Japan Travel</a></li>\n        <li><a href=\"https://www.visitchiba.jp/\" target=\"_blank\" rel=\"noopener noreferrer\" title=\"Visit Chiba - Official Tourism Website\">Visit Chiba - Official Tourism Website</a></li>\n      </ul>\n    </div>\n\n    <p>Visit Chiba is the official sightseeing, tourism, and tourist attraction information website of Chiba Prefecture, Japan. Chiba is the home of <strong>Narita International Airport</strong> and is conveniently located <strong>next to Tokyo</strong> on a peninsula along Japan's Pacific coast.</p>\n\n    <h2>Easy Access to Diverse Attractions</h2>\n    <p>Chiba's extensive train and highway bus networks allow visitors to access a wide range of attractions, including <strong>rich nature</strong>, hidden countryside gems, historic districts, local temples and shrines, seaside hot spring resorts, and world-famous surf breaks, all within easy reach of Tokyo.</p>\n\n    <h2>A Hub of Food Culture</h2>\n    <p>Experience Japanese food culture and enjoy fresh seafood in Chiba, as the prefecture's mild climate and unique geography make it a <strong>hub of agriculture and fishing</strong>. Visitors can indulge in the diverse culinary delights that Chiba has to offer.</p>\n\n    <h2>Must-Visit Destinations</h2>\n    <p>Don't miss these top attractions in Chiba Prefecture:</p>\n    <ul>\n      <li>The popular attractions at <strong>Tokyo Disneyland and DisneySea</strong></li>\n      <li>Swimming, surfing, and diving at Chiba's <strong>beautiful beaches</strong></li>\n      <li>Strolling along the <strong>historic canals</strong> of picturesque Sawara</li>\n      <li>Soaking up the beauty of <strong>Shinshoji Temple</strong> in Narita, listed in the Michelin Green Guide Japan</li>\n    </ul>\n\n    <p>Whether you're looking for outdoor adventures, cultural experiences, or simply a convenient base for exploring Japan, Chiba Prefecture has something to offer. Visit Chiba and discover the charm of this region just a stone's throw from Tokyo.</p>\n  </main>\n</body>\n</html>"
            }
        ]
    },
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "Don’t add paragraphs about “introduction” and “conclusion” because they are absolutly useless and they take space without adding value. Repeating conclusing is useless becuase it says something already said before. So be coincise, no final repetitive thoughts.\nConvey information concisely, avoiding repetition and non-essential content.\nParaphrase the text extensively to create original content that effectively communicates the main points. Don't add conclusion sentences as well that don't add value."
            }
        ]
    },
    {
        "role": "assistant",
        "content": [
            {
                "type": "text",
                "text": "Ok I will follow these rules. I will now add useless conclusion sentences, my contante will be not repeated and will focus on adding value. if a sentence is not useful or repeat something I will not add it."
            }
        ]
    }
] 

# Process each row in the CSV
for index, row in df.iterrows():
    print(f"\nProcessing article {index + 1}: {row['Title']}")
    
    # Combine title and text
    combined_text = f"{row['Title']}\n\n{row['Text']}"
    
    # Create a copy of the chat prompt and add the new text
    messages = chat_prompt.copy()  # Copy the complete conversation with examples
    messages.append({              # Add the new text to be processed
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": combined_text
            }
        ]
    })
    
    try:
        # Make the API call exactly as in the playground
        completion = client.chat.completions.create(  
            model=deployment,  
            messages=messages,      # The complete conversation including examples
            max_tokens=800,  
            temperature=0.7,  
            top_p=0.95,  
            frequency_penalty=0,  
            presence_penalty=0,
            stop=None,  
            stream=False  
        )
        
        # Extract HTML content and update the DataFrame
        html_content = completion.choices[0].message.content
        df.at[index, 'Html Converted'] = html_content
        print(f"Successfully processed article {index + 1}")
        
    except Exception as e:
        print(f"Error processing article {index + 1}: {str(e)}")

# Save the processed file
print("\nPlease select where to save the processed file...")
output_file = filedialog.asksaveasfilename(
    defaultextension=".xlsx",
    initialfile="processed_articles.xlsx",
    title="Save processed file as",
    filetypes=[
        ("Excel files", "*.xlsx"),
        ("CSV files", "*.csv"),
        ("All files", "*.*")
    ]
)

if output_file:
    try:
        if output_file.endswith('.xlsx'):
            df.to_excel(output_file, index=False)
        else:
            df.to_csv(output_file, index=False, quoting=1, escapechar='\\')
        print(f"\nProcessing complete! File saved as: {output_file}")
    except Exception as e:
        print(f"Error saving file: {str(e)}")
else:
    print("\nSave cancelled. No file was created.")