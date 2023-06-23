using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace MongoDbCRUD.Models
{
    public class ExcelData
    {
        [BsonId]
        public ObjectId Id { get; set; }

        [Required]
        [BsonElement("device_id")]
        public string device_id{ get; set; }

        [Required]
        [BsonElement("power_type")]
        public int power_type{ get; set; }

        [Required]
        [BsonElement("description")]
        public string description{ get; set; }

        [BsonElement("event_date")]
        public DateTime event_date{ get; set; }

        [BsonElement("added_on")]
        public DateTime added_on{ get; set; }
    }
}