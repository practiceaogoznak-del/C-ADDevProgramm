using AccessManager.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessManager.Models
{
    public class UserDbContext : DbContext
    {
        // Строка подключения к твоей БД
        private const string ConnectionString =
            "Data Source=06t-sql01;Initial Catalog=Portal;User ID=06portaluser;Password=Goznak202509;MultipleActiveResultSets=True;TrustServerCertificate=True";

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder
                    .UseSqlServer(ConnectionString, opt => opt.UseCompatibilityLevel(120)); // <= добавь это
            }
        }


        // Таблица UsersAD
        public DbSet<UserProfileEntity> UsersAD { get; set; }
    }
}
