namespace CanottaggioGui
{
    public class Athlete
    {
        public string Name { get; set; }
        public string Nation { get; set; }

        public override string ToString()
        {
            return $"{Name} ({Nation})";
        }
    }
}