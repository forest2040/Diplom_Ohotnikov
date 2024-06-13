using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Diplom_Ohotnikov.Tests
{
    [TestClass]
    public class PasswordDigitTest
    {
        [TestMethod]
        // Тестирование пароля на содержание цифр
        public void PasswordContainsNumbers()
        {
            string password = "password123";
            bool containsNumbers = ContainsNumbers(password);
            Assert.IsTrue(containsNumbers);
        }
        [TestMethod]
        private bool ContainsNumbers(string password)
        {
            foreach (char c in password)
            {
                if (Char.IsDigit(c))
                {
                    return true;
                }
            }
            return false;
        }
    }
}

