document.addEventListener("DOMContentLoaded", () => {
    const videoDataElement = document.getElementById("videoData");
    const subscriberCountElement = document.getElementById("subscriberCount");
    const loader = document.getElementById("loader");

    // Function to format duration
    function formatDuration(duration) {
        const match = duration.match(/PT(\d+H)?(\d+M)?(\d+S)?/);
        let formattedDuration = '';

        if (match[1]) {
            const hours = parseInt(match[1], 10);
            if (hours >= 1) {
                formattedDuration += hours + 'H ';
            }
        }

        if (match[2]) {
            const minutes = parseInt(match[2], 10);
            formattedDuration += minutes + 'M ';
        }

        if (match[3]) {
            const seconds = parseInt(match[3], 10);
            formattedDuration += seconds + 'S';
        }

        return formattedDuration.trim();
    }

    // Function to fetch video data from YouTube Data API
    function fetchVideoData() {
        loader.style.display = "block";
        videoDataElement.innerHTML = "";

        const apiKey = 'AIzaSyD0gbH6qSaSGJNhU4TsQH-Xs8genUcuGEc';
        const channelId = 'UCYY7eTARH5Ayk2UyVQSRhmQ';
        //const channelId = 'UC6Am-_05gR85sDkCwrsOa5A';

        const videoDataURL = `https://www.googleapis.com/youtube/v3/search?key=${apiKey}&channelId=${channelId}&order=viewCount&type=video&maxResults=100`;

        fetch(videoDataURL)
            .then(response => response.json())
            .then(data => {
                const videoIds = data.items.map(item => item.id.videoId);
                const videoDetailsPromises = videoIds.map(videoId => {
                    const videoDetailsURL = `https://www.googleapis.com/youtube/v3/videos?key=${apiKey}&id=${videoId}&part=snippet,statistics,contentDetails`;
                    return fetch(videoDetailsURL)
                        .then(response => response.json());
                });

                return Promise.all(videoDetailsPromises);
            })
            .then(videoDetails => {
                const formattedData = videoDetails.map(video => {
                    const snippet = video.items[0].snippet;
                    const statistics = video.items[0].statistics;
                    const contentDetails = video.items[0].contentDetails;
                    return {
                        title: snippet.title,
                        publishedAt: snippet.publishedAt,
                        views: statistics.viewCount,
                        likes: statistics.likeCount,
                        comments: statistics.commentCount,
                        averageViewDuration: formatDuration(contentDetails.duration),
                    };
                });

                formattedData.forEach(video => {
                    const row = document.createElement("tr");
                    row.innerHTML = `
                        <td>${video.title}</td>
                        <td>${video.publishedAt}</td>
                        <td>${video.views}</td>
                        <td>${video.likes}</td>
                        <td>${video.averageViewDuration}</td>
                        <td>${video.comments}</td>
                    `;
                    videoDataElement.appendChild(row);
                });
            })
            .catch(error => {
                console.error('Error fetching video data:', error);
            })
            .finally(() => {
                loader.style.display = "none";
            });
    }

    // Function to fetch subscriber count from YouTube Data API
    function fetchSubscriberCount() {
        const apiKey = 'AIzaSyD0gbH6qSaSGJNhU4TsQH-Xs8genUcuGEc';
        const channelId = 'UCYY7eTARH5Ayk2UyVQSRhmQ';
        //const channelId = 'UC6Am-_05gR85sDkCwrsOa5A';

        const subscriberCountURL = `https://www.googleapis.com/youtube/v3/channels?key=${apiKey}&id=${channelId}&part=statistics`;

        fetch(subscriberCountURL)
            .then(response => response.json())
            .then(data => {
                subscriberCountElement.textContent = data.items[0].statistics.subscriberCount;
            })
            .catch(error => {
                console.error('Error fetching subscriber count:', error);
            });
    }
    function downloadAsExcel() {
        const videoData = [];

        // Extract data from the table and store it in the videoData array
        const tableRows = videoDataElement.querySelectorAll("tr");
        tableRows.forEach(row => {
            const columns = row.querySelectorAll("td");
            if (columns.length === 6) {
                const rowData = [
                    columns[0].textContent,
                    columns[1].textContent,
                    columns[2].textContent,
                    columns[3].textContent,
                    columns[4].textContent,
                    columns[5].textContent
                ];
                videoData.push(rowData);
            }
        });

        // Create a new Excel workbook
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet([['Title', 'Published At', 'Views', 'Likes', 'Avg View Duration', 'Comments'], ...videoData]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Video Data');

        // Save the Excel file
        XLSX.writeFile(workbook, 'video_data.xlsx');
    }

    // Add click event listener to the "Download as Excel" button
    document.getElementById("downloadExcel").addEventListener("click", downloadAsExcel);

    // Fetch data when the page loads
    fetchVideoData();
    fetchSubscriberCount();
});
