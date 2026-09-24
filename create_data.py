import pandas as pd

data = {
    "Customer ID": [
        "C001", "C002", "C003", "C004", "C005",
        "C006", "C007", "C008", "C009", "C010",
        "C011", "C012", "C013", "C014", "C015",
        "C016", "C017", "C018", "C019", "C020"
    ],

    "Customer Name": [
        "John Smith", "Emma Wilson", "Michael Brown", "Sophia Davis",
        "Daniel Miller", "Olivia Taylor", "James Anderson", "Ava Thomas",
        "William Jackson", "Isabella White", "Robert Harris", "Mia Martin",
        "David Thompson", "Emily Garcia", "Chris Martinez", "Sarah Robinson",
        "Alex Clark", "Jessica Lewis", "Matthew Lee", "Grace Walker"
    ],

    "Email": [
        "john@gmail.com",
        "emma@gmail.com",
        "michael@gmail.com",
        "sophia@gmail.com",
        "daniel@gmail.com",
        "emma@gmail.com",
        "james@gmail.com",
        "ava@gmail.com",
        "william@gmail.com",
        "isabella@gmail.com",
        "robert@gmail.com",
        "mia@gmail.com",
        "david@gmail.com",
        "emily@gmail.com",
        "chris@gmail.com",
        "sarah@gmail.com",
        "alex@gmail.com",
        "jessica@gmail.com",
        "matthew@gmail.com",
        "john@gmail.com"
    ],

    "Country": [
        "USA",
        "United States",
        "UK",
        "United Kingdom",
        "Canada",
        "USA",
        "United States",
        "UK",
        "United Kingdom",
        "Canada",
        "USA",
        "United States",
        "UK",
        "United Kingdom",
        "Canada",
        "USA",
        "United States",
        "UK",
        "United Kingdom",
        "Canada"
    ],

    "Phone": [
        "+1 202-555-0101",
        "2025550102",
        "+44 20 7946 0958",
        "020-7946-0959",
        "+1 (202) 555-0105",
        "202-555-0106",
        "+1 202 555 0107",
        "020 7946 0961",
        "+44-20-7946-0962",
        "12345",
        "+1 202-555-0111",
        "2025550112",
        "+44 20 7946 0964",
        "02079460965",
        "+1-202-555-0116",
        "202-555-0117",
        "+1 202 555 0118",
        "020 7946 0969",
        "+44 20 7946 0970",
        "987654"
    ],

    "Order Value": [
        250, 450, 1200, 850, 600,
        300, 950, 1500, 700, 400,
        1100, 550, 800, 1250, 650,
        350, 900, 1350, 500, 750
    ]
}

df = pd.DataFrame(data)

df.to_excel("customer_data.xlsx", index=False)

print("customer_data.xlsx created successfully!")
print("\nTotal Records:", len(df))
print("\nData Preview:")
print(df.head())