namespace RootLogic.Models
{
    public static class Errors
    {
        public static CustomError NotInInterval = new CustomError("Not in interval", "Value is not in wanted interval.");
        public static CustomError InvalidLength = new CustomError("Invalid length", "Value does not have wanted length.");
        public static CustomError ValueNotAllowed = new CustomError("Value not allowed", "Value is not allowed.");
        public static CustomError OnlyNumber = new CustomError("Only numbers", "Cell should contain only numbers.");
    }
}
