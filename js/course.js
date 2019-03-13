// A class of course.
function Course(courseName) {
    /** @param courseName : String */
    this.name = courseName;
    this.category = new Set();  // A course may belong to multiple categories.
    this.freq = [];  // How many times the course appears in course1 to course24.
    for (var i = 0; i < 24; i++) { this.freq.push(0); }
}

Course.prototype.overallSum = function() {
    var sum = 0;
    for (var i = 0; i < this.freq.length; i++) { sum += this.freq[i]; }
    return sum;
};

Course.prototype.semesterSum = function() {
    var sum = [0, 0, 0, 0];
    for (var i = 0; i < this.freq.length; i++) {
        if (i >= 0 && i <= 5) { sum[0] += this.freq[i]; }
        if (i >= 6 && i <= 11) { sum[1] += this.freq[i]; }
        if (i >= 12 && i <= 17) { sum[2] += this.freq[i]; }
        if (i >= 18 && i <= 23) { sum[3] += this.freq[i]; }
    }
    return sum;
};