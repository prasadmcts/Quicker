using System;
using System.Collections.Generic;
using System.Linq;

namespace FirstTask
{
    class Program
    {
        static void Main(string[] args)
        {
            Solution sol = new Solution();
            Console.WriteLine($"For 201: {sol.solution(201)} ");
            Console.WriteLine($"For 202: {sol.solution(202)} ");
            Console.ReadLine();
        }
    }

    public class Solution
    {
        public string solution(int X)
        {
            string output = "";
            var data = LoadData();
            var dataTree = GetReleation(data);
            var category = dataTree.FirstOrDefault(t => t.CategoryId == X);
            if (category != null)
            {
                output = $"ParentCategoryID={category.ParentCategoryId}, Name={category.Name}";
                var keyWord = string.Join(",", data.Where(c => category.ChildCategories.Contains(c.CategoryId)).Select(c => c.Keywords));
                if (keyWord != null)
                {
                    output += $", Keywords={keyWord.TrimEnd(',')}.";
                }
            }
            return output;
        }

        public List<Category> GetReleation(List<Category> data)
        {
            data.ForEach(c =>
            {
                c.ChildCategories = RecursiveChildren(data, c.ParentCategoryId);
            });

            return data;
        }

        public List<int> RecursiveChildren(List<Category> categories, int parentCategoryId)
        {
            List<int> inner = new List<int>();
            foreach (var t in categories.Where(c => c.CategoryId == parentCategoryId).ToList())
            {
                inner.Add(t.CategoryId);
                if (t.ParentCategoryId != -1)
                {
                    inner = inner.Union(RecursiveChildren(categories, t.ParentCategoryId)).ToList();
                }
            }

            return inner;
        }

        private static List<Category> LoadData()
        {
            return new List<Category>()
            {
                new Category {CategoryId = 100, ParentCategoryId = -1, Name = "Business", Keywords="Money"},
                new Category {CategoryId = 200, ParentCategoryId = -1, Name = "Tutoring", Keywords="Teaching"},
                new Category {CategoryId = 101, ParentCategoryId = 100, Name = "Accouting", Keywords="Taxes"},
                new Category {CategoryId = 102, ParentCategoryId = 100, Name = "Taxation"},
                new Category {CategoryId = 201, ParentCategoryId = 200, Name = "Computer"},
                new Category {CategoryId = 103, ParentCategoryId = 101, Name = "Corporate Tax" },
                new Category {CategoryId = 202, ParentCategoryId = 201, Name = "Operationg System" },
                new Category {CategoryId = 109, ParentCategoryId = 101, Name = "Small business Tax" },
            };
        }
    }

    public class Category
    {
        public int CategoryId { get; set; }
        public int ParentCategoryId { get; set; }
        public string Name { get; set; }
        public string Keywords { get; set; }

        public List<int> ChildCategories { get; set; }
    }
}
