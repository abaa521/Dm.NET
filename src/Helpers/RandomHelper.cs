namespace Dm.NET.Helpers
{
    public class RandomHelper
    {
        private static readonly Random random = new();

        public static int GetRandomNumberMove()
        {
            return random.Next(0, 6); // 返回0到5的隨機整數
        }
    }
}