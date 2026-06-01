document.addEventListener('DOMContentLoaded', () => {
    const contentDiv = document.getElementById('content');
    
    // Fetch the markdown file
    fetch('Automation_Case_Study.md')
        .then(response => {
            if (!response.ok) {
                throw new Error('Failed to load the markdown file');
            }
            return response.text();
        })
        .then(markdown => {
            // Configure marked
            marked.setOptions({
                gfm: true,
                breaks: true
            });
            
            // Render markdown to HTML
            let htmlContent = marked.parse(markdown);
            
            // Post-process GitHub alerts for styling
            htmlContent = htmlContent.replace(/<blockquote>\s*<p>\[!WARNING\]/gi, '<blockquote class="warning"><p>');
            htmlContent = htmlContent.replace(/<blockquote>\s*<p>\[!NOTE\]/gi, '<blockquote class="note"><p>');
            htmlContent = htmlContent.replace(/<blockquote>\s*<p>\[!IMPORTANT\]/gi, '<blockquote class="warning"><p>');

            contentDiv.innerHTML = htmlContent;
            
            // Add slight fade-in animation
            contentDiv.style.opacity = 0;
            contentDiv.style.transition = 'opacity 0.6s ease-in';
            
            setTimeout(() => {
                contentDiv.style.opacity = 1;
            }, 50);
        })
        .catch(error => {
            console.error('Error loading markdown:', error);
            contentDiv.innerHTML = `
                <div style="text-align: center; padding: 40px;">
                    <h2>오류 발생</h2>
                    <p>문서를 불러오는 중 문제가 발생했습니다.</p>
                    <p style="color: #ff6b6b; font-size: 0.9em; margin-top: 10px;">${error.message}</p>
                </div>
            `;
        });
});
