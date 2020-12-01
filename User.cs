namespace ExpoIP
{
    public class User
    {
        /*these three fields a required for a user registration*/
        public string Email;
        public string FirstName;
        public string LastName;

        /*these two fields are optional and do not require to be set*/
        public string Title;
        /*the int will be assigned to a value by expoip*/
        /* 1 = Mr / 2 = Ms / 4 = Mx or other */
        public int Salutation;

        public User(string firstName, string lastName, string email, string title, int salutation) {

            FirstName = firstName;
            LastName = lastName;
            Email = email;
            Title = title;
            Salutation = salutation;

        }

    }
}
