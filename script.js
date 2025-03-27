// Global variables
let allData = [];
let calendarData = [];
let selectedView = 'yearly';
let selectedYear = '2023';
let selectedMonth = '1';
let selectedApp = 'all';

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    initializeControls();
    loadData();
});

// Set up event listeners for controls
function initializeControls() {
    // View tabs (yearly/monthly)
    document.querySelectorAll('.tab-btn[data-view]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn[data-view]').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            selectedView = btn.dataset.view;
            
            // Show/hide month select based on view
            document.getElementById('month-control').style.display = 
                selectedView === 'monthly' ? 'flex' : 'none';
            
            renderCalendar();
        });
    });
    
    // Application tabs
    document.querySelectorAll('.tab-btn[data-app]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn[data-app]').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            selectedApp = btn.dataset.app;
            renderCalendar();
        });
    });
    
    // Year select
    document.getElementById('year-select').addEventListener('change', (e) => {
        selectedYear = e.target.value;
        renderCalendar();
    });
    
    // Month select
    document.getElementById('month-select').addEventListener('change', (e) => {
        selectedMonth = e.target.value;
        renderCalendar();
    });
}

// Load and process Excel data
async function loadData() {
    try {
        const calendarContainer = document.getElementById('calendar');
        const loadingElement = document.getElementById('loading');
        
        // Show loading indicator
        loadingElement.style.display = 'block';
        
        // Load Excel files using fetch
        const [zoomResponse, firefoxResponse, webexResponse] = await Promise.all([
            fetch('./data/Zoom.xlsx').then(res => res.arrayBuffer()),
            fetch('./data/Firefox.xlsx').then(res => res.arrayBuffer()),
            fetch('./data/Webex.xlsx').then(res => res.arrayBuffer())
        ]);
        
        // Parse Excel files
        const zoomWorkbook = XLSX.read(new Uint8Array(zoomResponse), { cellDates: true });
        const firefoxWorkbook = XLSX.read(new Uint8Array(firefoxResponse), { cellDates: true });
        const webexWorkbook = XLSX.read(new Uint8Array(webexResponse), { cellDates: true });
        
        // Convert to JSON
        const zoomData = XLSX.utils.sheet_to_json(zoomWorkbook.Sheets[zoomWorkbook.SheetNames[0]]);
        const firefoxData = XLSX.utils.sheet_to_json(firefoxWorkbook.Sheets[firefoxWorkbook.SheetNames[0]]);
        const webexData = XLSX.utils.sheet_to_json(webexWorkbook.Sheets[webexWorkbook.SheetNames[0]]);
        
        // Add application name to each record
        const zoomWithApp = zoomData.map(item => ({...item, application: 'Zoom'}));
        const firefoxWithApp = firefoxData.map(item => ({...item, application: 'Firefox'}));
        const webexWithApp = webexData.map(item => ({...item, application: 'Webex'}));
        
        // Combine all data
        allData = [...zoomWithApp, ...firefoxWithApp, ...webexWithApp];
        
        // Process data for calendar
        processDataForCalendar();
        
        // Hide loading indicator
        loadingElement.style.display = 'none';
        
        // Render the calendar
        renderCalendar();
        
    } catch (error) {
        console.error('Error loading data:', error);
        document.getElementById('loading').textContent = 'Error loading data. Please check console for details.';
    }
}

// Process the data for calendar display
function processDataForCalendar() {
    // Group data by date
    const byDate = _.groupBy(allData, item => {
        const date = new Date(item['Release Date']);
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    });
    
    // Format for calendar heatmap
    calendarData = Object.entries(byDate).map(([date, items]) => {
        // Group by application
        const byApp = _.groupBy(items, 'application');
        
        // Get counts by application
        const zoomCount = (byApp['Zoom'] || []).length;
        const firefoxCount = (byApp['Firefox'] || []).length;
        const webexCount = (byApp['Webex'] || []).length;
        const totalCount = items.length;
        
        // Get Zoom categories if available
        const zoomCategories = zoomCount > 0 
            ? _.countBy(byApp['Zoom'], 'Group / Category')
            : {};
            
        // Split the date for convenience
        const [year, month, day] = date.split('-').map(d => parseInt(d));
        
        return {
            date,
            year,
            month,
            day,
            value: totalCount,
            zoom: zoomCount,
            firefox: firefoxCount,
            webex: webexCount,
            details: zoomCategories
        };
    });
}

// Main function to render the calendar based on selected options
function renderCalendar() {
    const calendarContainer = document.getElementById('calendar');
    calendarContainer.innerHTML = '';
    
    // Render statistics
    renderStats();
    
    if (selectedView === 'yearly') {
        renderYearView(calendarContainer, parseInt(selectedYear));
    } else {
        renderMonthView(calendarContainer, parseInt(selectedYear), parseInt(selectedMonth));
    }
}

// Render the summary statistics
function renderStats() {
    const statsContainer = document.getElementById('stats');
    statsContainer.innerHTML = '';
    
    // Get filtered data based on selection
    const filteredData = getFilteredData();
    
    if (filteredData.length === 0) {
        statsContainer.innerHTML = '<p class="no-data">No data available for the selected period</p>';
        return;
    }
    
    // Calculate statistics
    const totalFeatures = filteredData.reduce((sum, day) => sum + day.value, 0);
    const zoomFeatures = filteredData.reduce((sum, day) => sum + day.zoom, 0);
    const firefoxFeatures = filteredData.reduce((sum, day) => sum + day.firefox, 0);
    const webexFeatures = filteredData.reduce((sum, day) => sum + day.webex, 0);
    
    // Create statistics cards
    statsContainer.innerHTML = `
        <div class="stat-card">
            <div class="stat-value">${totalFeatures}</div>
            <div class="stat-label">Total Features</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">
                <span class="stat-icon zoom-color"></span>
                ${zoomFeatures}
            </div>
            <div class="stat-label">Zoom Features</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">
                <span class="stat-icon firefox-color"></span>
                ${firefoxFeatures}
            </div>
            <div class="stat-label">Firefox Features</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">
                <span class="stat-icon webex-color"></span>
                ${webexFeatures}
            </div>
            <div class="stat-label">Webex Features</div>
        </div>
    `;
}

// Render the yearly calendar view (all months)
function renderYearView(container, year) {
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    const yearlyContainer = document.createElement('div');
    yearlyContainer.className = 'yearly-calendar';
    
    // Create a container for each month
    months.forEach((monthName, index) => {
        const monthNumber = index + 1;
        
        const monthContainer = document.createElement('div');
        monthContainer.className = 'month-container';
        
        // Month header
        const monthHeader = document.createElement('div');
        monthHeader.className = 'month-header';
        monthHeader.textContent = monthName;
        monthContainer.appendChild(monthHeader);
        
        // Weekday headers
        const weekdaysContainer = document.createElement('div');
        weekdaysContainer.className = 'weekdays';
        ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach(day => {
            const dayEl = document.createElement('div');
            dayEl.textContent = day;
            weekdaysContainer.appendChild(dayEl);
        });
        monthContainer.appendChild(weekdaysContainer);
        
        // Calendar grid
        const calendarGrid = document.createElement('div');
        calendarGrid.className = 'calendar-grid';
        
        // Add month's days
        generateMonthDays(calendarGrid, year, monthNumber);
        
        monthContainer.appendChild(calendarGrid);
        yearlyContainer.appendChild(monthContainer);
    });
    
    container.appendChild(yearlyContainer);
}

// Render the monthly calendar view (single month)
function renderMonthView(container, year, month) {
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    const monthlyContainer = document.createElement('div');
    monthlyContainer.className = 'monthly-calendar';
    
    // Month container
    const monthContainer = document.createElement('div');
    monthContainer.className = 'month-container';
    
    // Month header
    const monthHeader = document.createElement('div');
    monthHeader.className = 'month-header';
    monthHeader.textContent = `${months[month-1]} ${year}`;
    monthContainer.appendChild(monthHeader);
    
    // Weekday headers
    const weekdaysContainer = document.createElement('div');
    weekdaysContainer.className = 'weekdays';
    ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'].forEach(day => {
        const dayEl = document.createElement('div');
        dayEl.textContent = day;
        weekdaysContainer.appendChild(dayEl);
    });
    monthContainer.appendChild(weekdaysContainer);
    
    // Calendar grid
    const calendarGrid = document.createElement('div');
    calendarGrid.className = 'calendar-grid';
    
    // Add month's days
    generateMonthDays(calendarGrid, year, month);
    
    monthContainer.appendChild(calendarGrid);
    monthlyContainer.appendChild(monthContainer);
    container.appendChild(monthlyContainer);
}

// Generate days for a specific month
function generateMonthDays(container, year, month) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const firstDayOfMonth = new Date(year, month - 1, 1).getDay();
    
    // Add empty cells for days before the 1st of the month
    for (let i = 0; i < firstDayOfMonth; i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = 'calendar-day empty';
        container.appendChild(emptyCell);
    }
    
    // Add cells for each day of the month
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayData = calendarData.find(d => d.date === dateStr);
        
        const dayCell = document.createElement('div');
        dayCell.className = 'calendar-day';
        
        // Add day number
        const dayNumber = document.createElement('div');
        dayNumber.className = 'day-number';
        dayNumber.textContent = day;
        dayCell.appendChild(dayNumber);
        
        if (dayData) {
            const value = getValueForApp(dayData);
            
            // Set cell background color based on value (heatmap)
            if (value > 0) {
                const opacity = Math.min(0.1 + (value / 10) * 0.9, 1);
                const blueValue = Math.floor(255 - opacity * 200);
                dayCell.style.backgroundColor = `rgb(235, 245, ${blueValue})`;
                
                // Create bubbles container
                const bubblesContainer = document.createElement('div');
                bubblesContainer.className = 'bubbles-container';
                
                // Add bubble for each application with features
                if (dayData.zoom > 0 && (selectedApp === 'all' || selectedApp === 'zoom')) {
                    addBubble(bubblesContainer, dayData.zoom, 'zoom-color');
                }
                
                if (dayData.firefox > 0 && (selectedApp === 'all' || selectedApp === 'firefox')) {
                    addBubble(bubblesContainer, dayData.firefox, 'firefox-color');
                }
                
                if (dayData.webex > 0 && (selectedApp === 'all' || selectedApp === 'webex')) {
                    addBubble(bubblesContainer, dayData.webex, 'webex-color');
                }
                
                dayCell.appendChild(bubblesContainer);
                
                // Add title with feature count
                dayCell.title = `${dateStr}: ${dayData.value} features`;
            }
        }
        
        container.appendChild(dayCell);
    }
}

// Add a bubble to the container
function addBubble(container, count, colorClass) {
    const bubble = document.createElement('div');
    bubble.className = `bubble ${colorClass}`;
    
    // Calculate bubble size based on count (min 20px, max 40px)
    const size = Math.max(20, Math.min(40, count * 5));
    bubble.style.width = `${size}px`;
    bubble.style.height = `${size}px`;
    
    // Add count inside bubble if large enough
    if (count > 3) {
        bubble.textContent = count;
    }
    
    container.appendChild(bubble);
}

// Get filtered data based on selected year and month
function getFilteredData() {
    if (!calendarData.length) return [];
    
    let filtered = [...calendarData];
    
    // Filter by year
    if (selectedYear) {
        filtered = filtered.filter(item => item.year === parseInt(selectedYear));
    }
    
    // Filter by month if in monthly view
    if (selectedView === 'monthly' && selectedMonth) {
        filtered = filtered.filter(item => item.month === parseInt(selectedMonth));
    }
    
    return filtered;
}

// Get value based on selected application
function getValueForApp(item) {
    if (selectedApp === 'all') return item.value;
    if (selectedApp === 'zoom') return item.zoom;
    if (selectedApp === 'firefox') return item.firefox;
    if (selectedApp === 'webex') return item.webex;
    return 0;
}