import React, { useState, useEffect } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  CircularProgress,
  Alert,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  List,
  ListItem,
  ListItemText,
  ListItemIcon,
  Divider,
  Chip,
  Stack,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Grid
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import WarningIcon from '@mui/icons-material/Warning';
import TrendingUpIcon from '@mui/icons-material/TrendingUp';
import GroupIcon from '@mui/icons-material/Group';
import LightbulbIcon from '@mui/icons-material/Lightbulb';
import axios from 'axios';

function SprintAnalysis() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [analysis, setAnalysis] = useState(null);
  const [boards, setBoards] = useState([]);
  const [sprints, setSprints] = useState([]);
  const [selectedBoard, setSelectedBoard] = useState('');
  const [selectedSprint, setSelectedSprint] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [analysisResult, setAnalysisResult] = useState(null);
  const [selectedFile, setSelectedFile] = useState(null);
  const [isDownloading, setIsDownloading] = useState(false);

  useEffect(() => {
    fetchBoards();
  }, []);

  useEffect(() => {
    if (selectedBoard) {
      fetchSprints(selectedBoard);
    } else {
      setSprints([]);
      setSelectedSprint('');
    }
  }, [selectedBoard]);

  const fetchBoards = async () => {
    try {
      const response = await axios.get('http://localhost:5000/api/boards');
      setBoards(response.data);
      if (response.data.length > 0) {
        setSelectedBoard(response.data[0].id);
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to fetch boards');
    }
  };

  const fetchSprints = async (boardId) => {
    try {
      const response = await axios.get(`http://localhost:5000/api/sprints?boardId=${boardId}`);
      setSprints(response.data);
      if (response.data.length > 0) {
        setSelectedSprint(response.data[0].id);
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to fetch sprints');
    }
  };

  const handleFileChange = (event) => {
    const selectedFile = event.target.files[0];
    if (selectedFile && selectedFile.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      setFile(selectedFile);
      setError(null);
    } else {
      setError('Please select a valid Excel file (.xlsx)');
      setFile(null);
    }
  };

  const handleUpload = async () => {
    if (!file) {
      setError('Please select a file first');
      return;
    }

    if (!selectedBoard || !selectedSprint) {
      setError('Please select both a board and a sprint');
      return;
    }

    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('boardId', selectedBoard);
    formData.append('sprintId', selectedSprint);

    try {
      const response = await axios.post('http://localhost:5000/api/sprint-analysis', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
      setAnalysis(response.data);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to analyze sprint data');
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    if (!analysis) return;
    
    setIsDownloading(true);
    try {
      const response = await fetch('http://localhost:5000/api/sprint-analysis/download', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          sprint_data: {
            sprint_name: analysis.sprint_name,
            sprint_goal: analysis.sprint_goal,
            start_date: analysis.start_date,
            end_date: analysis.end_date
          },
          improvement_areas: analysis.improvement_areas
        })
      });

      if (!response.ok) {
        throw new Error('Failed to download document');
      }

      // Get the blob from the response
      const blob = await response.blob();
      
      // Create a download link
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `sprint_analysis_${analysis.sprint_name}.docx`;
      document.body.appendChild(a);
      a.click();
      
      // Clean up
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (error) {
      console.error('Error downloading document:', error);
      setError('Failed to download document. Please try again.');
    } finally {
      setIsDownloading(false);
    }
  };

  const renderSpillOverAnalysis = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        Spilled Stories
      </Typography>
      <List>
        {data.spilled_stories.map((story, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <WarningIcon color="error" />
            </ListItemIcon>
            <ListItemText
              primary={story.story_id}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Reason: {story.reason}
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Prevention: {story.prevention_suggestion}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Root Causes
      </Typography>
      <List>
        {data.root_causes.map((cause, index) => (
          <ListItem key={index}>
            <ListItemText primary={cause} />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Recommendations
      </Typography>
      <List>
        {data.recommendations.map((rec, index) => (
          <ListItem key={index}>
            <ListItemText primary={rec} />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const renderChurnAnalysis = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        High Churn Stories
      </Typography>
      <List>
        {data.high_churn_stories.map((story, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <TrendingUpIcon color="warning" />
            </ListItemIcon>
            <ListItemText
              primary={story.story_id}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Churn Count: {story.churn_count}
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Impact: {story.impact}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Velocity Impact
      </Typography>
      <Typography paragraph>{data.velocity_impact}</Typography>
      <Typography variant="h6" gutterBottom>
        Reduction Suggestions
      </Typography>
      <List>
        {data.reduction_suggestions.map((suggestion, index) => (
          <ListItem key={index}>
            <ListItemText primary={suggestion} />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const renderTeamUtilization = (data) => (
    <Box>
      <Typography variant="h6" gutterBottom>
        Under-Utilized Team Members
      </Typography>
      <List>
        {data.under_utilized.map((member, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <GroupIcon color="info" />
            </ListItemIcon>
            <ListItemText
              primary={member.member}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Utilization: {member.utilization}%
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Suggestion: {member.suggestion}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Over-Utilized Team Members
      </Typography>
      <List>
        {data.over_utilized.map((member, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <GroupIcon color="error" />
            </ListItemIcon>
            <ListItemText
              primary={member.member}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Utilization: {member.utilization}%
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Suggestion: {member.suggestion}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
      <Typography variant="h6" gutterBottom sx={{ mt: 2 }}>
        Workload Distribution
      </Typography>
      <Typography paragraph>{data.workload_distribution}</Typography>
      <Typography variant="h6" gutterBottom>
        Optimization Suggestions
      </Typography>
      <List>
        {data.optimization_suggestions.map((suggestion, index) => (
          <ListItem key={index}>
            <ListItemText primary={suggestion} />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const renderAdditionalImprovements = (data) => (
    <Box>
      <List>
        {data.map((improvement, index) => (
          <ListItem key={index}>
            <ListItemIcon>
              <LightbulbIcon color="primary" />
            </ListItemIcon>
            <ListItemText
              primary={improvement.area}
              secondary={
                <>
                  <Typography component="span" variant="body2" color="text.primary">
                    Observation: {improvement.observation}
                  </Typography>
                  <br />
                  <Typography component="span" variant="body2" color="text.secondary">
                    Suggestion: {improvement.suggestion}
                  </Typography>
                </>
              }
            />
          </ListItem>
        ))}
      </List>
    </Box>
  );

  return (
    <Box sx={{ p: 3 }}>
      <Paper elevation={3} sx={{ p: 3, mb: 3 }}>
        <Typography variant="h5" gutterBottom>
          Sprint Analysis
        </Typography>
        <Typography variant="body1" paragraph>
          Select a board and sprint, then upload your sprint capacity Excel file to analyze improvement areas, including spill-over analysis,
          churn analysis, and team utilization.
        </Typography>

        <Grid container spacing={2} sx={{ mb: 3 }}>
          <Grid item xs={12} md={6}>
            <FormControl fullWidth>
              <InputLabel id="board-select-label">Select Board</InputLabel>
              <Select
                labelId="board-select-label"
                id="board-select"
                value={selectedBoard}
                label="Select Board"
                onChange={(e) => setSelectedBoard(e.target.value)}
              >
                {boards.map((board) => (
                  <MenuItem key={board.id} value={board.id}>
                    {board.name}
                  </MenuItem>
                ))}
              </Select>
            </FormControl>
          </Grid>
          <Grid item xs={12} md={6}>
            <FormControl fullWidth>
              <InputLabel id="sprint-select-label">Select Sprint</InputLabel>
              <Select
                labelId="sprint-select-label"
                id="sprint-select"
                value={selectedSprint}
                label="Select Sprint"
                onChange={(e) => setSelectedSprint(e.target.value)}
                disabled={!selectedBoard}
              >
                {sprints.map((sprint) => (
                  <MenuItem key={sprint.id} value={sprint.id}>
                    {sprint.name} ({new Date(sprint.startDate).toLocaleDateString()} - {new Date(sprint.endDate).toLocaleDateString()})
                  </MenuItem>
                ))}
              </Select>
            </FormControl>
          </Grid>
        </Grid>
        
        <Box sx={{ mb: 2 }}>
          <input
            accept=".xlsx"
            style={{ display: 'none' }}
            id="excel-file-upload"
            type="file"
            onChange={handleFileChange}
          />
          <label htmlFor="excel-file-upload">
            <Button
              variant="outlined"
              component="span"
              sx={{ mr: 2 }}
            >
              Select Excel File
            </Button>
          </label>
          {file && (
            <Chip
              label={file.name}
              onDelete={() => setFile(null)}
              sx={{ ml: 2 }}
            />
          )}
        </Box>

        <Button
          variant="contained"
          onClick={handleUpload}
          disabled={!file || loading || !selectedBoard || !selectedSprint}
          sx={{ mt: 2 }}
        >
          {loading ? <CircularProgress size={24} /> : 'Analyze Sprint'}
        </Button>

        {error && (
          <Alert severity="error" sx={{ mt: 2 }}>
            {error}
          </Alert>
        )}
      </Paper>

      {analysis && (
        <Paper elevation={3} sx={{ p: 3 }}>
          <Typography variant="h5" gutterBottom>
            Analysis Results
          </Typography>

          <Accordion>
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography variant="h6">Spill-over Analysis</Typography>
            </AccordionSummary>
            <AccordionDetails>
              {renderSpillOverAnalysis(analysis.improvement_areas.spill_over_analysis)}
            </AccordionDetails>
          </Accordion>

          <Accordion>
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography variant="h6">Churn Analysis</Typography>
            </AccordionSummary>
            <AccordionDetails>
              {renderChurnAnalysis(analysis.improvement_areas.churn_analysis)}
            </AccordionDetails>
          </Accordion>

          <Accordion>
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography variant="h6">Team Utilization</Typography>
            </AccordionSummary>
            <AccordionDetails>
              {renderTeamUtilization(analysis.improvement_areas.team_utilization)}
            </AccordionDetails>
          </Accordion>

          <Accordion>
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Typography variant="h6">Additional Improvements</Typography>
            </AccordionSummary>
            <AccordionDetails>
              {renderAdditionalImprovements(analysis.improvement_areas.additional_improvements)}
            </AccordionDetails>
          </Accordion>

          <div className="analysis-header">
            <h3>Sprint Analysis Results</h3>
            <button 
              className="download-button"
              onClick={handleDownload}
              disabled={isDownloading}
            >
              {isDownloading ? 'Downloading...' : 'Download Report'}
            </button>
          </div>
        </Paper>
      )}
    </Box>
  );
}

export default SprintAnalysis; 