using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Diplom_Ohotnikov.Tests
{
    [TestClass]
    public class PasswordLengthTest
    {
        [TestMethod]
        // Тестирование длины пароля
        public void TestPasswordLength()
        {
            string password = "password";
            int length = password.Length;
            Assert.IsTrue(length >= 8);
        }
    }
}

