// 等待页面完全渲染
setTimeout(() => {
    // 尝试多种可能的父元素（兼容不同Swagger版本）
    const possibleParents = [
        document.querySelector('.topbar-wrapper'),
        document.querySelector('.swagger-ui .header'),
        document.querySelector('.swagger-ui > div:first-child'),
        document.body
    ];

    // 创建链接
    const hangfireLink = document.createElement('a');
    hangfireLink.href = '/hangfire'; // 确认Hangfire路径正确
    hangfireLink.target = '_blank';
    hangfireLink.textContent = '数据同步面板';
    hangfireLink.className = 'hangfire-link';

    // 尝试添加到第一个可用的父元素
    for (const parent of possibleParents) {
        if (parent) {
            parent.appendChild(hangfireLink);
            break;
        }
    }
}, 100); // 延迟1秒，确保Swagger UI完全加载