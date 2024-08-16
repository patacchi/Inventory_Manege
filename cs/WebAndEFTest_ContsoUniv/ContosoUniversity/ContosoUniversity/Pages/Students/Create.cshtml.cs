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
        //View Model ���g�p���鎞�͂�����
        [BindProperty]
        public StudentVM StudentVM { get; set; }
        public async Task<IActionResult> OnPostAsync()
        {
            if(!ModelState.IsValid)
            {
                return Page();
            }
            var entry = _context.Add(new Student());
            entry.CurrentValues.SetValues(StudentVM);   //SetValues�ł͂��̃I�u�W�F�N�g�̒l��ݒ肷��̂ɕʂ̃I�u�W�F�N�g����l��ǂݎ��
                                                        //Student() �̐ݒ������̂� StudentVM�̃f�[�^���g�p�����(�v���p�e�B������)
            await _context.SaveChangesAsync();
            return RedirectToPage("./Index");
        }

        //TryUpdateModelAsync���g�p����ꍇ�͂�����
        /*
        [BindProperty]
        public Student Student { get; set; }
        public async Task<IActionResult> OnPostAsync()
        {
            var emptystudent = new Student();
            if(await TryUpdateModelAsync<Student>(
                emptystudent,
                "student",  //Prefix for form value.���̃v���t�B�b�N�X��t���ăt�H�[���t�B�[���h������
                            //�Ⴆ�� student.FIirstName �t�B�[���h�Ƃ��A�啶���������͋�ʂ���Ȃ�
                s => s.FirstMidName, s => s.LastName, s => s.EnrollmentDate)) //���X�g���ꂽ�l�̂ݎg�p
            {
                _context.Students.Add(emptystudent);
                await _context.SaveChangesAsync();
                return RedirectToPage("./Index");
            }
        
            /*���̃R�[�h�Aoverpost�U���ɑ΂��ĐƎ�
            _context.Students.Add(Student);     �ݒ肷��v���p�e�B�����肵�Ă��Ȃ�
            await _context.SaveChangesAsync();

            return RedirectToPage("./Index");
            
            return Page();
        }*/
    }
}
