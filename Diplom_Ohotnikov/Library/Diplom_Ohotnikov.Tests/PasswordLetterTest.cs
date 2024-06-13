using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Diplom_Ohotnikov.Tests
{
    [TestClass]
    public class PasswordLetterTest
    {
        [TestMethod]
        // Тестирование пароля на содержание заглавной буквы
        public void PasswordContainsUppercase()
        {
            string password = "Password123";
            bool containsUppercase = PasswordValidator.ContainsUppercaseLetter(password);
            Assert.IsTrue(containsUppercase);
        }

        public static class PasswordValidator
        {
            public static bool ContainsUppercaseLetter(string password)
            {
                foreach (char c in password)
                {
                    if (char.IsUpper(c))
                    {
                        return true;
                    }
                }

                return false;
            }
        }
    }
}