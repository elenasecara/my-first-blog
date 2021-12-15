import xlsxwriter

from django.shortcuts import redirect, render, get_object_or_404
from django.utils import timezone
from .models import Post
from .forms import PostForm

# Create your views here.
def post_list(request):
    posts = Post.objects.filter(published_date__lte=timezone.now()).order_by('published_date')
    return render(request, 'blog/post_list.html', {'posts': posts})

def post_detail(request, pk):
    post = get_object_or_404(Post, pk=pk)
    return render(request, 'blog/post_detail.html', {'post': post})

def post_new(request):
    if request.method == "POST":
        form = PostForm(request.POST)
        if form.is_valid():
            post = form.save(commit=False)
            post.author = request.user
            post.published_date = timezone.now()
            post.save()
            try:
                write_to_spreadsheet()
            except:
                return render(request, 'blog/post_edit.html', {'form': form})
            return redirect('post_detail', pk=post.pk)
    else:
        form = PostForm
    return render(request, 'blog/post_edit.html', {'form': form})

def post_edit(request, pk):
    post = get_object_or_404(Post, pk=pk)
    if(request.method == "POST"):
        form = PostForm(request.POST, instance=post)
        if form.is_valid():
            post = form.save(commit=False)
            post.author = request.user
            post.published_date = timezone.now()
            post.save()
            return redirect('post_detail', pk=post.pk)
    else:
        form = PostForm(instance=post)
    return render(request, 'blog/post_edit.html', {'form': form})

def write_to_spreadsheet():
    print("trying to write to spreadsheet")
    # the file to be created
    workbook = xlsxwriter.Workbook('C:/Users/sgb584/Desktop/booking.xlsx')

    # add new worksheet
    worksheet = workbook.add_worksheet()

    row = 0
    column = 0

    content = ["ankit", "rahul", "priya", "harshita",
                    "sumit", "neeraj", "shivam"]

    for item in content:
        worksheet.write(row, column, item)
        row += 1
    workbook.close()