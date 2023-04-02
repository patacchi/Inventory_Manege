using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using ContosoUniversity.Data;
using ContosoUniversity.Models;

namespace ContosoUniversity.pages_Students
{
    public class CreateModel : PageModel
    {
        private readonly ContosoUniversity.Data.SchoolContext _context;

        public CreateModel(ContosoUniversity.Data.SchoolContext context)
        {
            _context = context;
        }

        public IActionResult OnGet()
        {
            return Page();
        }
        // To protect from overposting attacks, see https://aka.ms/RazorPagesCRUD
        //View Model を使用する時はこちら
        [BindProperty]
        public StudentVM StudentVM { get; set; }
        public async Task<IActionResult> OnPostAsync()
        {
            if(!ModelState.IsValid)
            {
                return Page();
            }
            var entry = _context.Add(new Student());
            entry.CurrentValues.SetValues(StudentVM);   //SetValuesではこのオブジェクトの値を設定するのに別のオブジェクトから値を読み取る
                                                        //Student() の設定をするのに StudentVMのデータが使用される(プロパティ名同一)
            await _context.SaveChangesAsync();
            return RedirectToPage("./Index");
        }

        //TryUpdateModelAsyncを使用する場合はこちら
        /*
        [BindProperty]
        public Student Student { get; set; }
        public async Task<IActionResult> OnPostAsync()
        {
            var emptystudent = new Student();
            if(await TryUpdateModelAsync<Student>(
                emptystudent,
                "student",  //Prefix for form value.このプリフィックスを付けてフォームフィールドを検索
                            //例えば student.FIirstName フィールドとか、大文字小文字は区別されない
                s => s.FirstMidName, s => s.LastName, s => s.EnrollmentDate)) //リストされた値のみ使用
            {
                _context.Students.Add(emptystudent);
                await _context.SaveChangesAsync();
                return RedirectToPage("./Index");
            }
        
            /*元のコード、overpost攻撃に対して脆弱
            _context.Students.Add(Student);     設定するプロパティを限定していない
            await _context.SaveChangesAsync();

            return RedirectToPage("./Index");
            
            return Page();
        }*/
    }
}
