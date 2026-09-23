#!/usr/bin/env bash
# QuestionNinja - MongoDB Setup Script
# Database: questionninja
# Collection: papers

set -e

MONGO_URI="${MONGO_URI:-mongodb://localhost:27017/questionninja}"

echo "Connecting to MongoDB at: $MONGO_URI"

mongosh "$MONGO_URI" --eval '
const colName = "papers";

if (!db.getCollectionNames().includes(colName)) {
  print("Creating collection: " + colName);
  db.createCollection(colName, {
    validator: {
      $jsonSchema: {
        bsonType: "object",
        required: ["branding", "metadata", "sections", "updatedAt"],
        properties: {
          title: { bsonType: "string" },
          branding: { bsonType: "object" },
          metadata: { bsonType: "object" },
          sections: { bsonType: "array" },
          createdAt: { bsonType: "date" },
          updatedAt: { bsonType: "date" }
        }
      }
    }
  });
} else {
  print("Collection already exists: " + colName);
}

// Ensure indexes
db.papers.createIndex({ "metadata.subject": 1 });
db.papers.createIndex({ "updatedAt": -1 });

print("Current total documents in " + colName + ": " + db.papers.countDocuments());
print("MongoDB database and collection are ready!");
'
