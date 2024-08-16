namespace ContosoUniversity.Models
{
    public class Student
    {
        public int ID { get; set; }
        public string LastName { get; set; }
        public string FirstMidName { get; set; }
        public DateTime EnrollmentDate { get; set; }

        public ICollection<Enrollment> Enrollments { get; set; }
    }
    /// <summary>
    /// ビューモデルの定義 一部のフィールドのみ必要なページ等で使用する
    /// モデルの種類に関連付ける必要はないが、プロパティ名は同じにする必要がある
    /// </summary>
    public class StudentVM
    {
        public int ID { get; set; }
        public string LastName { get; set; }
        public string FirstMidName { get; set; }
        public DateTime EnrollmentDate { get; set; }
    }
}