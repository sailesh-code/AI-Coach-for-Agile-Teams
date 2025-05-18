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
  Grid,
  Card,
  CardContent,
  Stepper,
  Step,
  StepLabel
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import WarningIcon from '@mui/icons-material/Warning';
import TrendingUpIcon from '@mui/icons-material/TrendingUp';
import GroupIcon from '@mui/icons-material/Group';
import LightbulbIcon from '@mui/icons-material/Lightbulb';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import DescriptionIcon from '@mui/icons-material/Description';
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
  const [activeStep, setActiveStep] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [analysisResult, setAnalysisResult] = useState(null);
  const [selectedFile, setSelectedFile] = useState(null);
  const [isDownloading, setIsDownloading] = useState(false);

  const steps = ['Select Board & Sprint', 'Upload Excel File', 'Generate Report'];

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
      setActiveStep(1);
    } else {
      setError('Please select a valid Excel file (.xlsx)');
      setFile(null);
    }
  };

  const handleGenerateReport = async () => {
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
    setActiveStep(2);

    const formData = new FormData();
    formData.append('file', file);
    formData.append('boardId', selectedBoard);
    formData.append('sprintId', selectedSprint);

    try {
      const response = await axios.post('http://localhost:5000/api/sprint-combined-report', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob'
      });

      // Create a download link
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `sprint_report_and_analysis_${selectedSprint}.docx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      
      // Reset the form
      setFile(null);
      setActiveStep(0);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to generate report');
      setActiveStep(1);
    } finally {
      setLoading(false);
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
          Sprint Report & Analysis
        </Typography>
        
        <Stepper activeStep={activeStep} sx={{ mb: 4 }}>
          {steps.map((label) => (
            <Step key={label}>
              <StepLabel>{label}</StepLabel>
            </Step>
          ))}
        </Stepper>

        <Grid container spacing={3}>
          <Grid item xs={12} md={6}>
            <Card>
              <CardContent>
                <Typography variant="h6" gutterBottom>
                  Board & Sprint Selection
                </Typography>
                <FormControl fullWidth sx={{ mb: 2 }}>
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
              </CardContent>
            </Card>
          </Grid>
          
          <Grid item xs={12} md={6}>
            <Card>
              <CardContent>
                <Typography variant="h6" gutterBottom>
                  Excel File Upload
                </Typography>
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
                    startIcon={<CloudUploadIcon />}
                    fullWidth
                    sx={{ mb: 2 }}
                  >
                    Select Excel File
                  </Button>
                </label>
                {file && (
                  <Chip
                    label={file.name}
                    onDelete={() => {
                      setFile(null);
                      setActiveStep(0);
                    }}
                    sx={{ width: '100%' }}
                  />
                )}
              </CardContent>
            </Card>
          </Grid>
        </Grid>

        <Box sx={{ mt: 3, textAlign: 'center' }}>
          <Button
            variant="contained"
            onClick={handleGenerateReport}
            disabled={!file || loading || !selectedBoard || !selectedSprint}
            startIcon={loading ? <CircularProgress size={20} /> : <DescriptionIcon />}
            size="large"
          >
            {loading ? 'Generating Report...' : 'Generate Report'}
          </Button>
        </Box>

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
        </Paper>
      )}
    </Box>
  );
}

export default SprintAnalysis; 