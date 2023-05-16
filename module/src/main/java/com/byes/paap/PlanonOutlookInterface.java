package com.byes.paap;

import java.text.SimpleDateFormat;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.models.BodyType;
import com.microsoft.graph.models.DateTimeTimeZone;
import com.microsoft.graph.models.Event;
import com.microsoft.graph.models.FreeBusyStatus;
import com.microsoft.graph.models.ItemBody;
import com.microsoft.graph.models.Location;
import com.microsoft.graph.requests.GraphServiceClient;
import com.planonsoftware.platform.backend.businessrule.v3.IBusinessRule;
import com.planonsoftware.platform.backend.businessrule.v3.IBusinessRuleContext;
import com.planonsoftware.platform.backend.data.v1.IBusinessObject;

public class PlanonOutlookInterface implements IBusinessRule {
    @Override
    public void execute(IBusinessObject newBO, IBusinessObject oldBO, IBusinessRuleContext context) {

        if (newBO.getStateName().equals("Assigned") && !oldBO.getStateName().equals("Assigned")) {
            // BYES
            String clientId = "bb44f77f-eb36-4742-afc0-91c1844c24d3";
            String clientSecret = "cQMaD_-5C_wkw0QiCsA.cFvn_2Fwa4d3DH";
            String tenant = "4a3d9983-e936-4837-9552-9d9126a92eb0";
            
            IBusinessObject person = newBO.getReferenceFieldByName("ResourcePersonRef").getValue();
            String email = person.getStringFieldByName("Email").getValueAsString();

            if (email == null) {
                return;
            }

            IBusinessObject order = newBO.getReferenceFieldByName("OrderRef").getValue();

            IBusinessObject property = order.getReferenceFieldByName("PropertyRef").getValue();
            String propertyText = property.getStringFieldByName("Name").getValueAsString();

            Date startDate = newBO.getDateTimePropertyFieldByName("PlannedBeginDatetime").getValueAsDateTime();
            Date endDate = newBO.getDateTimePropertyFieldByName("PlannedEndDatetime").getValueAsDateTime();
            
            if (startDate == null || endDate == null) {
                return;
            }
            
            SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss");
            String startDateString = formatter.format(startDate);
            String endDateString = formatter.format(endDate);
            
            SimpleDateFormat formatter2 = new SimpleDateFormat("dd.MM.yyyy HH:mm");
            String clientStartDateString = formatter2.format(startDate);
            String clientEndDateString = formatter2.format(endDate);

            IBusinessObject requester = order.getReferenceFieldByName("InternalRequestorPersonRef").getValue();
            String requesterName = "";
            if (requester != null) {
                requesterName = requester.getStringFieldByName("LastName").getValueAsString() + ", " +requester.getStringFieldByName("FirstName").getValueAsString();
            }

            String propertyCode = property.getStringFieldByName("Code").getValueAsString();
            String propertyName = property.getStringFieldByName("Name").getValueAsString();
            String propertyAddress = property.getStringFieldByName("Address").getValueAsString();
            IBusinessObject propertyPLZ = property.getReferenceFieldByName("FreeString9").getValue();
            propertyAddress = propertyAddress + ", " + propertyPLZ.getStringFieldByName("FreeString1").getValueAsString();
            propertyAddress = propertyAddress + propertyPLZ.getStringFieldByName("Name").getValueAsString();

            IBusinessObject customer = order.getReferenceFieldByName("CustomerRef").getValue();
            String customerCode = customer.getStringFieldByName("Code").getValueAsString();
            String customerName = customer.getStringFieldByName("Name").getValueAsString();

            IBusinessObject wbsItem = order.getReferenceFieldByName("FreeInteger2").getValue();
            String wbsItemString = "";
            if (wbsItem != null) {
                wbsItemString = wbsItem.getStringFieldByName("Code").getValueAsString() + ", " + wbsItem.getStringFieldByName("Name").getValueAsString();
            }
            String orderComments = "";
            if (order.getStringFieldByName("Comment").getValueAsString() != null) {
                orderComments = order.getStringFieldByName("Comment").getValueAsString();
            }
            String orderTitle = order.getStringFieldByName("OrderNumber").getValueAsString() + "-" + order.getStringFieldByName("Description").getValueAsString();
            String orderText = "<html><body><table>";
            orderText = orderText + "<tr>Kundentermin Start:<td></td><td>" + clientStartDateString + "</td></tr>";
            orderText = orderText + "<tr>Kundentermin Ende:<td></td><td>" + clientEndDateString + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Auftragsnummer:<td></td><td>" + order.getStringFieldByName("OrderNumber").getValueAsString()  + "</td></tr>";
            orderText = orderText + "<tr>Bezeichnung:<td></td><td>" + order.getStringFieldByName("Description").getValueAsString() + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Beschreibung:<td></td><td>" + orderComments + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Melder:<td></td><td>" + requesterName + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Objekt:<td></td><td>" + propertyCode + "<br>" + propertyName + ", " + propertyAddress + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Ansprechpartner vor Ort:<td></td><td></td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Kunde:<td></td><td>" + customerCode + "<br>" + customerName + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>PSP-Element:<td></td><td>" + wbsItemString + "</td></tr>";
            orderText = orderText + "<tr><td></td><td></td></tr>";
            orderText = orderText + "<tr>Disponent:<td></td><td></td></tr>";
            orderText = orderText + "</table></body></html>";

            final ClientSecretCredential clientSecretCredential = new ClientSecretCredentialBuilder()
                    .clientId(clientId)
                    .clientSecret(clientSecret)
                    .tenantId(tenant)
                    .build();
                    
            List<String> scopes = new ArrayList<String>();
            scopes.add("https://graph.microsoft.com/.default");

            final TokenCredentialAuthProvider tokenCredentialAuthProvider = new TokenCredentialAuthProvider(scopes, clientSecretCredential);

            final GraphServiceClient graphClient = GraphServiceClient.builder().authenticationProvider(tokenCredentialAuthProvider).buildClient();
                
            Event event = new Event();
            event.subject = orderTitle;
            ItemBody body = new ItemBody();
            body.contentType = BodyType.HTML;
            body.content = orderText;
            event.body = body;
            DateTimeTimeZone start = new DateTimeTimeZone();
            start.dateTime = startDateString;
            start.timeZone = "Europe/Berlin";
            event.start = start;
            DateTimeTimeZone end = new DateTimeTimeZone();
            end.dateTime = endDateString;
            end.timeZone = "Europe/Berlin";
            event.end = end;
            Location location = new Location();
            location.displayName = propertyText;
            event.location = location;
            event.showAs = FreeBusyStatus.TENTATIVE;
            
            event.allowNewTimeProposals = true;
            long timestamp = Instant.now().getEpochSecond();
            event.transactionId = "WA_" + newBO.getPrimaryKeyAsString() + "_" + timestamp;
            
            graphClient.users(email).events().buildRequest().post(event);
        }
        
    }
}