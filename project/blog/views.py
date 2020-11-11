from django.shortcuts import render, get_object_or_404
from .models import Post
import markdown


# Create your views here.


def index(request):
    post = Post.objects.all().order_by('-created_time')
    post = post[0:3]
    return render(request, 'blog/index.html', {
        'title': 'Excel -> CSV直行便',
        'post': post,
    })


def detail(request, pk):
    post = get_object_or_404(Post, pk=pk)
    post.body = markdown.markdown(post.body,
                                  extensions=[
                                      'markdown.extensions.extra',
                                      'markdown.extensions.codehilite',
                                      'markdown.extensions.toc',
                                  ])
    return render(request, 'blog/detail.html', {
        'post': post,
        'title': post.title
    })
